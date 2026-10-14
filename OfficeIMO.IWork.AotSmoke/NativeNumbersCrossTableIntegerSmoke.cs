using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersCrossTableIntegerSmoke {
    internal static void Run() {
        Check("integer-cross-table.numbers", false);
        Check("integer-cross-table-changed.numbers", true);
        Console.WriteLine("PASS | native Numbers cross-table integer calls | two tables/two sheets, 8 numeric GCD/LCM calls, arithmetic control, native caches, saved operand recalculation and explicit padding loss");
    }

    private static void Check(string name, bool changed) {
        string path = Path.Combine(AppContext.BaseDirectory, name);
        string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
        Require(hash == (changed ? "0DBFFD11551FFF2D830A671F04E0E7FE0F6E071156F50E2E98F2D31C4E1835BD"
            : "5551AAD4E3BA7E2B918681C39282499A6E409FDE7139C268DE32C63367DF68FD"),
            "Native cross-table integer fixture changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var tables = result.Projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        Require(tables.Select(table => table.Name).SequenceEqual(new[] { "Data", "Totals", "Empty" })
            && tables.Sum(table => table.Cells.Count(cell => cell.Kind == IWorkCellKind.Formula && cell.FormulaIsComplete)) == 9,
            "Native cross-table formulas or table order changed.");
        Require(result.Report.IsPartialEditableReconstruction
            && result.Report.Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED")
            && !result.Report.Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_FORMULA_PARTIAL"),
            "Native cross-table formula or padding diagnostics changed.");
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        Verify(reopened, changed);
        Require(reopened.Calculate() == 9, "Cross-table calculation count changed."); Verify(reopened, changed);
        reopened.Sheets[0].CellValue(2, 2, 24d);
        Require(reopened.Calculate() == 9, "Edited cross-table calculation count changed."); Verify(reopened, true);
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes); Verify(final, true);
        Require(!final.ValidateOpenXml().Any()
            && Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash,
            "Cross-table XLSX validation or input immutability failed.");
    }

    private static void Verify(ExcelDocument document, bool changed) {
        Require(document.Sheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "Budget - Data", "Budget - Totals", "Empty grid" }),
            "Cross-table destination names changed.");
        (int Sheet, int Row, int Column, string Formula, double Value)[] cells = {
            (0, 2, 4, "GCD('Budget - Totals'!B1,6)", 3d),
            (0, 3, 4, "LCM('Budget - Totals'!B1,6)", 198d), (0, 4, 4, "30+6", 36d),
            (2, 1, 1, "GCD('Budget - Data'!B2:B4)", changed ? 12d : 6d),
            (2, 1, 2, "LCM('Budget - Data'!B2:B4)", changed ? 360d : 180d),
            (2, 1, 3, "GCD('Budget - Data'!B2,'Budget - Totals'!B1)", 3d),
            (2, 2, 1, "LCM('Budget - Data'!$B$2,'Budget - Totals'!$B$1)", changed ? 792d : 990d),
            (2, 2, 2, "GCD('Budget - Data'!B2:B4,'Budget - Totals'!B1)", 3d),
            (2, 2, 3, "LCM('Budget - Data'!B2,'Budget - Totals'!B1,0)", 0d)
        };
        foreach (var cell in cells) {
            var sheet = document.Sheets[cell.Sheet];
            Require(sheet.GetFormulaText(cell.Row, cell.Column) == cell.Formula
                && sheet.CellAt(cell.Row, cell.Column).GetValue<double>() == cell.Value,
                "Cross-table formula text, native cache or recalculated value changed.");
        }
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
