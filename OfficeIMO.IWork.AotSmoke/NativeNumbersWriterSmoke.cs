using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersWriterSmoke {
    private const string AcceptedSha256 = "08F6F8C2B82B99514D657FDBB0402E3FAD25EF7C7005C9B84AFAFD929D1E719C";

    internal static void Run() {
        var document = IWorkNumbersDocument.Create();
        var data = IWorkTableData.Create(4, 4).SetText(1, 1, "Item").SetText(1, 2, "Value").SetText(1, 3, "Flag").SetText(1, 4, "Notes")
            .SetText(2, 1, "Apples").SetNumber(2, 2, 12.5).SetBoolean(2, 3, true).SetText(2, 4, "Zażółć 😀")
            .SetText(3, 1, "Negative").SetNumber(3, 2, -0.125).SetBoolean(3, 3, false).SetText(3, 4, "=literal")
            .SetText(4, 1, "Line1\nLine2").SetNumber(4, 2, 12345.6789).SetBoolean(4, 3, true);
        var first = document.AddSheet("Budget");
        first.AddTable(data, "Data", 40, 80, 130, 36, "Arial", 14, "003366");
        first.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "Totals").SetNumber(1, 2, 99).SetBoolean(2, 1, false),
            "Totals", 40, 300, 130, 36);
        document.AddSheet("Empty grid").AddTable(IWorkTableData.Create(2, 3), "Empty", 30, 40, 80, 28);
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedSha256 && bytes.Length == 22_090,
            "The NativeAOT Numbers creator differs from the independently decoded Apple-accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Numbers creation is not deterministic.");
        using var destination = new MemoryStream(); destination.WriteByte(9); document.Save(destination);
        Require(destination.CanWrite && bytes.SequenceEqual(destination.ToArray().Skip(1)), "Numbers stream save differs or closes its caller.");
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(bytes).ReadNumbers();
        Require(projection.Sheets.Count == 2 && projection.Sheets[0].Tables.Count == 2
            && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0
            && projection.Diagnostics.Count == 0, "Numbers creation lost its complete sheet/table graph.");
        IWorkTable table = projection.Sheets[0].Tables[0], empty = projection.Sheets[1].Tables[0];
        Require((double)table.GetCell(2, 2)!.Value! == 12.5 && (bool)table.GetCell(2, 3)!.Value!
            && !(bool)table.GetCell(3, 3)!.Value! && (string)table.GetCell(2, 4)!.Value! == "Zażółć 😀"
            && (string)table.GetCell(4, 1)!.Value! == "Line1\nLine2"
            && table.GetCell(3, 4)!.Kind == IWorkCellKind.Text && table.RowHeights[1] == 36
            && table.ColumnWidths[1] == 130 && empty.RowCount == 2 && empty.ColumnCount == 3 && empty.Cells.Count == 0,
            "Numbers scalar types, Unicode, dimensions or empty positions differ.");
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Require(converted.Report.IsPartialEditableReconstruction
            && converted.Report.Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED")
            && empty.GetPadding(2, 3)?.BottomPoints == 4, "Native Numbers region padding was silently discarded.");
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Require(reopened.Sheets.Count == 3 && reopened.Sheets[0].CellAt(2, 2).GetValue<double>() == 12.5
            && reopened.Sheets[0].CellAt(2, 3).GetValue<bool>() && !reopened.Sheets[0].CellAt(3, 3).GetValue<bool>()
            && reopened.Sheets[0].CellAt(3, 4).GetValue<string>() == "=literal"
            && string.IsNullOrEmpty(reopened.Sheets[0].GetFormulaText(3, 4))
            && !reopened.ValidateOpenXml().Any(), "Saved Numbers XLSX loses typed scalar values or literal text.");
        byte[] before = destination.ToArray();
        try {
            document.Save(destination, new() { MaximumTableCells = 25 });
            throw new InvalidOperationException("Numbers dense-grid limit was not enforced.");
        } catch (InvalidDataException) {
            Require(before.SequenceEqual(destination.ToArray()), "Numbers budget rejection mutated its destination.");
        }
        Console.WriteLine($"PASS | native Numbers scalar creation | {AcceptedSha256} | {bytes.Length} bytes | deterministic sheets/tables, Decimal128, Unicode, booleans, empty grid, caller stream, saved/reopened XLSX, explicit padding omission and dense-grid bounds");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
