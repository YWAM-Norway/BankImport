using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace BankImport
{
    public sealed class BankTransaction
    {
        public string UtfortDato { get; set; } = string.Empty;
        public string BokførtDato { get; set; } = string.Empty;
        public string Rentedato { get; set; } = string.Empty;
        public string Beskrivelse { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Undertype { get; set; } = string.Empty;
        public string FraKonto { get; set; } = string.Empty;
        public string Avsender { get; set; } = string.Empty;
        public string TilKonto { get; set; } = string.Empty;
        public string Mottakernavn { get; set; } = string.Empty;
        public string BeløpInn { get; set; } = string.Empty;
        public string BeløpUt { get; set; } = string.Empty;
        public string Valuta { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Numref { get; set; } = string.Empty;
        public string Arkivref { get; set; } = string.Empty;
        public string MeldingKID { get; set; } = string.Empty;
    }

    public static class Camt053Parser
    {
        public static List<BankTransaction> Parse(string xmlPath)
        {
            if (string.IsNullOrWhiteSpace(xmlPath))
                throw new ArgumentException("XML file path is required.", nameof(xmlPath));

            XDocument document = XDocument.Load(xmlPath);
            XElement root = document.Root ?? throw new InvalidDataException("The XML document has no root element.");
            XNamespace primaryNamespace = root.Name.Namespace;
            XElement? addtlTxInf = root.Descendants().FirstOrDefault(e => e.Name.LocalName == "AddtlTxInf");
            XNamespace? secondaryNamespace = addtlTxInf?.Name.Namespace;
            XNamespace[] namespaces = secondaryNamespace == null || secondaryNamespace == primaryNamespace
                ? new[] { primaryNamespace }
                : new[] { primaryNamespace, secondaryNamespace };

            var rows = new List<BankTransaction>();
            foreach (XElement statement in root.Descendants().Where(e => e.Name.LocalName == "Stmt"))
            {
                string statementIban = ReadAccountId(Child(statement, "Acct", namespaces));
                foreach (XElement entry in statement.Descendants().Where(e => e.Name.LocalName == "Ntry"))
                {
                    // Do not accidentally consume an entry nested inside another entry.
                    XElement? enclosingEntry = entry.Ancestors().FirstOrDefault(a => a.Name.LocalName == "Ntry");
                    if (enclosingEntry != null)
                        continue;

                    List<XElement> details = entry.Descendants().Where(e => e.Name.LocalName == "TxDtls").ToList();
                    if (details.Count == 0)
                        rows.Add(CreateRow(entry, null, statementIban, namespaces));
                    else
                        foreach (XElement detail in details)
                            rows.Add(CreateRow(entry, detail, statementIban, namespaces));
                }
            }

            // Some files can expose Ntry outside a Stmt wrapper. Preserve those entries too.
            foreach (XElement entry in root.Descendants().Where(e => e.Name.LocalName == "Ntry" && !e.Ancestors().Any(a => a.Name.LocalName == "Stmt")))
            {
                if (entry.Ancestors().Any(a => a.Name.LocalName == "Ntry"))
                    continue;
                List<XElement> details = entry.Descendants().Where(e => e.Name.LocalName == "TxDtls").ToList();
                if (details.Count == 0)
                    rows.Add(CreateRow(entry, null, string.Empty, namespaces));
                else
                    foreach (XElement detail in details)
                        rows.Add(CreateRow(entry, detail, string.Empty, namespaces));
            }
            return rows;
        }

        private static BankTransaction CreateRow(XElement entry, XElement? detail, string statementIban, XNamespace[] namespaces)
        {
            string direction = Text(Child(entry, "CdtDbtInd", namespaces)).ToUpperInvariant();
            string counterpartySide = direction == "DBIT" ? "Cdtr" : "Dbtr";
            string otherSide = direction == "DBIT" ? "Dbtr" : "Cdtr";
            string counterpartyName = ReadPartyName(detail, counterpartySide, namespaces);
            string otherName = ReadPartyName(detail, otherSide, namespaces);

            List<string> remittance = detail == null
                ? new List<string>()
                : detail.Descendants().Where(e => e.Name.LocalName == "Ustrd")
                    .Select(Text).Where(s => s.Length > 0).ToList();
            string endToEnd = detail == null ? string.Empty : Text(FindPath(detail, namespaces, "Refs", "EndToEndId"));
            if (endToEnd.Length == 0 && detail != null)
                endToEnd = Text(detail.Descendants().FirstOrDefault(e => e.Name.LocalName == "EndToEndId"));

            XElement? amountElement = Child(entry, "Amt", namespaces);
            string rawAmount = Text(amountElement);
            string amountString = NormalizeAmount(rawAmount, direction);
            decimal amount = 0m;
            decimal.TryParse(amountString, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);

            string fromAccount;
            string toAccount;
            if (direction == "DBIT")
            {
                fromAccount = NonEmpty(statementIban, ReadRelatedAccount(detail, "DbtrAcct", namespaces));
                toAccount = ReadRelatedAccount(detail, "CdtrAcct", namespaces);
            }
            else
            {
                toAccount = NonEmpty(statementIban, ReadRelatedAccount(detail, "CdtrAcct", namespaces));
                fromAccount = ReadRelatedAccount(detail, "DbtrAcct", namespaces);
            }

            string archiveReference = Text(Child(entry, "AcctSvcrRef", namespaces));
            if (archiveReference.Length == 0)
                archiveReference = Text(Child(entry, "CcyXchgRef", namespaces));
            if (archiveReference.Length == 0)
            {
                XElement? ntryRef = entry.Descendants().FirstOrDefault(e => e.Name.LocalName == "NtryRef");
                archiveReference = Text(ntryRef);
            }
            if (archiveReference.Length == 0)
            {
                XElement? ntx = entry.Descendants().FirstOrDefault(e => e.Name.LocalName == "Ntx");
                archiveReference = Text(Child(ntx, "NtryRef", namespaces));
            }

            return new BankTransaction
            {
                UtfortDato = ReadDate(Child(entry, "BookgDt", namespaces), namespaces),
                BokførtDato = ReadDate(Child(entry, "ValDt", namespaces), namespaces),
                Rentedato = ReadDate(Child(entry, "RlDt", namespaces), namespaces),
                Beskrivelse = detail == null ? string.Empty : FirstRemittance(remittance),
                Type = counterpartyName,
                Undertype = FirstRemittance(remittance),
                FraKonto = fromAccount,
                Avsender = otherName,
                TilKonto = toAccount,
                Mottakernavn = counterpartyName,
                BeløpInn = amount > 0m ? amountString : string.Empty,
                BeløpUt = amount < 0m ? Math.Abs(amount).ToString("0.############################", CultureInfo.InvariantCulture) : string.Empty,
                Valuta = amountElement?.Attribute("Ccy")?.Value?.Trim() ?? string.Empty,
                Status = Text(Child(entry, "Sts", namespaces)),
                Numref = Text(Child(entry, "Nref", namespaces)),
                Arkivref = archiveReference,
                MeldingKID = remittance.Count > 0 ? string.Join(" ", remittance) : endToEnd
            };
        }

        private static string ReadDate(XElement? dateElement, XNamespace[] namespaces)
        {
            if (dateElement == null) return string.Empty;
            foreach (string name in new[] { "Dt", "Tm", "DtTm" })
            {
                XElement? value = Child(dateElement, name, namespaces);
                if (value != null && Text(value).Length > 0) return Text(value);
            }
            return Text(dateElement);
        }

        private static string ReadPartyName(XElement? detail, string side, XNamespace[] namespaces)
        {
            if (detail == null) return string.Empty;
            XElement? party = Child(detail, side + "Pty", namespaces) ?? Child(detail, side, namespaces);
            string name = Text(Child(party, "Nm", namespaces));
            if (name.Length > 0) return name;
            XElement? fallback = detail.Descendants().FirstOrDefault(e => e.Name.LocalName == side + "Pty")
                ?? detail.Descendants().FirstOrDefault(e => e.Name.LocalName == side);
            return Text(Child(fallback, "Nm", namespaces));
        }

        private static string ReadRelatedAccount(XElement? detail, string accountName, XNamespace[] namespaces)
        {
            if (detail == null) return string.Empty;
            XElement? relatedParties = detail.Descendants().FirstOrDefault(e => e.Name.LocalName == "RltdPties");
            XElement? account = Child(relatedParties, accountName, namespaces);
            if (account == null)
                account = detail.Descendants().FirstOrDefault(e => e.Name.LocalName == accountName);
            return ReadAccountId(account);
        }

        private static string ReadAccountId(XElement? account)
        {
            if (account == null) return string.Empty;
            XElement? iban = account.DescendantsAndSelf().FirstOrDefault(e => e.Name.LocalName == "IBAN");
            if (iban != null && Text(iban).Length > 0) return Text(iban);
            // GetInnerId-style fallback: use the value of the last descendant Id.
            XElement? lastId = account.Descendants().Where(e => e.Name.LocalName == "Id").LastOrDefault();
            return Text(lastId);
        }

        private static XElement? FindPath(XElement parent, XNamespace[] namespaces, params string[] path)
        {
            XElement? current = parent;
            foreach (string part in path)
            {
                current = Child(current, part, namespaces);
                if (current == null) return null;
            }
            return current;
        }

        private static XElement? Child(XElement? parent, string localName, XNamespace[] namespaces)
        {
            if (parent == null) return null;
            foreach (XNamespace ns in namespaces)
            {
                XElement? match = parent.Elements(ns + localName).FirstOrDefault();
                if (match != null) return match;
            }
            return parent.Elements().FirstOrDefault(e => e.Name.LocalName == localName);
        }

        private static string Text(XElement? element) => element?.Value.Trim() ?? string.Empty;
        private static string FirstRemittance(List<string> remittance) => remittance.Count > 0 ? remittance[0] : string.Empty;
        private static string NonEmpty(string first, string fallback) => first.Length > 0 ? first : fallback;

        private static string NormalizeAmount(string raw, string direction)
        {
            if (!decimal.TryParse(raw.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal parsed))
                return string.Empty;
            if (direction == "DBIT") parsed = -Math.Abs(parsed);
            else if (direction == "CRDT") parsed = Math.Abs(parsed);
            return parsed.ToString("0.############################", CultureInfo.InvariantCulture);
        }
    }

    public static class KontoMapping
    {
        public static readonly Dictionary<string, string> BankTilKonto = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["28014852737"] = "1920", ["30005365399"] = "1921", ["30003997371"] = "1945", ["30001469378"] = "1960",
            ["30001503665"] = "1979", ["30001469394"] = "1980", ["30001627763"] = "1990", ["30001469343"] = "1946",
            ["30001544647"] = "1961", ["30004911004"] = "1997", ["28007492415"] = "2241", ["30006758076"] = "2247",
            ["30002848852"] = "1947", ["30001499056"] = "1962", ["30001505846"] = "1991", ["28015163684"] = "1993",
            ["28013635481"] = "2242", ["30001469947"] = "1948", ["30004297942"] = "1963", ["30006441552"] = "2243",
            ["28014662083"] = "1949", ["28015075254"] = "1964", ["30001551686"] = "1950", ["30001551643"] = "1965",
            ["30005024168"] = "1998", ["28014065913"] = "1951", ["30002554298"] = "1952", ["30002554336"] = "1967",
            ["28007513021"] = "2244", ["30000739916"] = "1953", ["30000760672"] = "1968", ["30000740019"] = "1986",
            ["30000740043"] = "1994", ["30008457996"] = "2245", ["30002758705"] = "1954", ["30002758721"] = "1969",
            ["30002876929"] = "1988", ["30005066375"] = "1943", ["30005063953"] = "1944", ["30002119621"] = "1955",
            ["28015241421"] = "1970", ["30005944836"] = "2246", ["30003997363"] = "1957", ["28014995589"] = "1958",
            ["28014995570"] = "1971", ["30003997355"] = "1959", ["30003997347"] = "1972", ["30001397164"] = "1989",
            ["30006441560"] = "1973", ["30001469483"] = "1981", ["28008065618"] = "2240", ["28008123855"] = "2248"
        };

        public static readonly Dictionary<string, string> KontoTilBank = BankTilKonto
            .ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.Ordinal);

        public static string? FinnKonto(BankTransaction t)
        {
            if (t == null) return null;
            foreach (string raw in new[] { t.FraKonto, t.TilKonto })
            {
                string normalized = NormalizeAccount(raw);
                if (normalized.Length == 0) continue;
                if (BankTilKonto.TryGetValue(normalized, out string? account)) return account;
                if (normalized.StartsWith("NO", StringComparison.OrdinalIgnoreCase))
                    normalized = normalized.Substring(2);
                if (normalized.Length >= 11)
                    normalized = normalized.Substring(normalized.Length - 11);
                if (BankTilKonto.TryGetValue(normalized, out account)) return account;
            }
            return null;
        }

        public static string? FinnBank(string konto)
        {
            string normalized = (konto ?? string.Empty).Trim();
            return KontoTilBank.TryGetValue(normalized, out string? bank) ? bank : null;
        }

        private static string NormalizeAccount(string? value) => (value ?? string.Empty)
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("-", string.Empty, StringComparison.Ordinal)
            .Trim();
    }

    public static class CsvWriter
    {
        public static void Write(List<BankTransaction> rows, string path)
        {
            using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
            writer.WriteLine(string.Join(";", Headers.Select(Escape)));
            foreach (BankTransaction row in rows)
                writer.WriteLine(string.Join(";", Values(row).Select(Escape)));
        }

        internal static readonly string[] Headers =
        {
            "Utført dato", "Bokført dato", "Rentedato", "Beskrivelse", "Type", "Undertype", "Fra konto", "Avsender",
            "Til konto", "Mottakernavn", "Beløp inn", "Beløp ut", "Valuta", "Status", "Numref", "Arkivref", "Melding/KID/Fakt.nr"
        };

        internal static string[] Values(BankTransaction row) => new[]
        {
            row.UtfortDato, row.BokførtDato, row.Rentedato, row.Beskrivelse, row.Type, row.Undertype, row.FraKonto,
            row.Avsender, row.TilKonto, row.Mottakernavn, row.BeløpInn, row.BeløpUt, row.Valuta, row.Status,
            row.Numref, row.Arkivref, row.MeldingKID
        };

        private static string Escape(string? value)
        {
            string text = value ?? string.Empty;
            return text.IndexOfAny(new[] { ';', '"', '\r', '\n' }) >= 0
                ? "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""
                : text;
        }
    }

public static class ExcelExporter
{
    private static readonly CultureInfo NorwegianCulture = CultureInfo.GetCultureInfo("nb-NO");

    public static void Export(List<BankTransaction> rows, string path)
    {
        using var workbook = new XLWorkbook();
        IXLWorksheet worksheet = workbook.Worksheets.Add("Transactions");
        for (int column = 0; column < CsvWriter.Headers.Length; column++)
            worksheet.Cell(1, column + 1).Value = CsvWriter.Headers[column];

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            string[] values = CsvWriter.Values(rows[rowIndex]);
            for (int column = 0; column < values.Length; column++)
            {
                IXLCell cell = worksheet.Cell(rowIndex + 2, column + 1);
                if ((column == 10 || column == 11) &&
                    decimal.TryParse(values[column], NumberStyles.Number, NorwegianCulture, out decimal numeric))
                    cell.Value = (double)numeric;
                else
                    cell.Value = values[column] ?? string.Empty;
            }
        }
        worksheet.Row(1).Style.Font.Bold = true;
        worksheet.SheetView.FreezeRows(1);
        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(path);
    }
}
    public static class Program
    {
        public static int Main(string[] args)
        {
            try
            {
                if (args.Length != 1)
                    throw new ArgumentException("Usage: BankImport <camt.053 XML file>");

                string xmlPath = Path.GetFullPath(args[0]);
                List<BankTransaction> rows = Camt053Parser.Parse(xmlPath);
                string outputDirectory = Path.Combine(Path.GetDirectoryName(xmlPath) ?? Directory.GetCurrentDirectory(), "Kontoutskrifter");
                Directory.CreateDirectory(outputDirectory);

                var groups = rows.Select(row => new { Row = row, Konto = KontoMapping.FinnKonto(row) })
                    .GroupBy(item => item.Konto ?? "0000", StringComparer.Ordinal)
                    .OrderBy(group => group.Key, StringComparer.Ordinal)
                    .ToList();
                foreach (var group in groups)
                    CsvWriter.Write(group.Select(item => item.Row).ToList(), Path.Combine(outputDirectory, group.Key + ".csv"));

                // Always create the unmapped export, including when there are no unmapped rows.
                if (!groups.Any(group => group.Key == "0000"))
                    CsvWriter.Write(new List<BankTransaction>(), Path.Combine(outputDirectory, "0000.csv"));

                CsvWriter.Write(rows, Path.Combine(outputDirectory, "UIOTrans.csv"));
                ExcelExporter.Export(rows, Path.Combine(outputDirectory, "UIOTrans.xlsx"));

                foreach (var group in groups)
                    Console.WriteLine($"{group.Key}: {group.Count()} transaksjoner");
                int unmappedCount = rows.Count(row => KontoMapping.FinnKonto(row) == null);
                Console.WriteLine($"Umappede linjer skrevet til 0000.csv: {unmappedCount}");
                foreach (BankTransaction row in rows.Where(row => KontoMapping.FinnKonto(row) == null).Take(5))
                    Console.WriteLine($"Eksempel: FraKonto='{row.FraKonto}', TilKonto='{row.TilKonto}', beløp inn='{row.BeløpInn}', beløp ut='{row.BeløpUt}'");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception.Message);
                return 1;
            }
        }
    }
}
