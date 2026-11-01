using ClosedXML.Excel;
using System.Xml.Linq;

// ---------- Modeller ----------
public record BankTransaction(
    string Account,
    DateTime BookingDate,
    decimal Amount,
    string Currency,
    string Name,
    string Description);

public static class Camt053Parser
{
    public static List<BankTransaction> ParseFile(string path)
    {
        var doc = XDocument.Load(path);
        var ns = doc.Root!.GetDefaultNamespace();
        var txs = new List<BankTransaction>();

        foreach (var acct in doc.Descendants(ns + "Acct"))
        {
            var account = acct.Element(ns + "Id")?.Element(ns + "IBAN")?.Value
                       ?? acct.Element(ns + "Id")?.Element(ns + "Othr")?.Element(ns + "Id")?.Value
                       ?? "?";

            // Bal ligger som søsken til Ntry inne i Stmt
            foreach (var entry in acct.Parent!.Descendants(ns + "Ntry"))
            {
                var amtEl = entry.Element(ns + "Amt");
                if (amtEl == null) continue;

                var amount = decimal.Parse(amtEl.Value, System.Globalization.CultureInfo.InvariantCulture);
                if (entry.Element(ns + "CdtDbtInd")?.Value == "DBIT") amount = -amount;

                var dateStr = entry.Element(ns + "BookgDt")?.Element(ns + "Dt")?.Value;
                DateTime.TryParse(dateStr, out var date);

                var txDtls = entry.Descendants(ns + "TxDtls").FirstOrDefault();
                var name = txDtls?.Element(ns + "RltdPties")?.Element(ns + "Nm")?.Value ?? "";
                var desc = txDtls?.Element(ns + "AddtlNtryInf")?.Value
                        ?? entry.Descendants(ns + "Ustrd").FirstOrDefault()?.Value ?? "";

                txs.Add(new BankTransaction(account, date, amount,
                    amtEl.Attribute("Ccy")?.Value ?? "NOK", name, desc));
            }
        }
        return txs;
    }
}

// ---------- Excel-eksport ----------
public static class ExcelExporter
{
    public static void Write(List<BankTransaction> txs, string path)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Transaksjoner");

        // Overskrifter
        string[] headers = { "Konto", "Dato", "Beløp", "Valuta", "Motpart", "Beskrivelse" };
        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int row = 2;
        foreach (var t in txs)
        {
            ws.Cell(row, 1).Value = t.Account;
            ws.Cell(row, 2).Value = t.BookingDate;
            ws.Cell(row, 2).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 3).Value = t.Amount;
            ws.Cell(row, 3).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 4).Value = t.Currency;
            ws.Cell(row, 5).Value = t.Name;
            ws.Cell(row, 6).Value = t.Description;
            row++;
        }

        // Gjør området til en "tabell" → innebygd filtrering/søk i Excel
        var range = ws.Range(1, 1, row - 1, headers.Length);
        var table = ws.Tables.Add(range, "Transaksjoner");
        table.Theme = XLTableTheme.TableStyleMedium2;
        ws.Columns().AdjustToContents();

        wb.SaveAs(path);
    }
}

// ---------- Main ----------
var xmlPath = args.Length > 0 ? args[0] : @"C:\Users\DinBruker\Downloads\UIO Trans.xml";
var xlsxPath = Path.ChangeExtension(xmlPath, ".xlsx");

var txs = Camt053Parser.ParseFile(xmlPath);
ExcelExporter.Write(txs, xlsxPath);

Console.WriteLine($"{txs.Count} transaksjoner skrevet til {xlsxPath}");
foreach (var g in txs.GroupBy(t => t.Account))
    Console.WriteLine($"  {g.Key}: {g.Count()} transaksjoner, sum {g.Sum(t => t.Amount):N2}");