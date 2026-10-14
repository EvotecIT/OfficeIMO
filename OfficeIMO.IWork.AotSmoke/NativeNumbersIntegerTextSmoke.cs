using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersIntegerTextSmoke {
    internal static void Run() {
        Verify("integer-text-reference.numbers", "c7af6b7727f6dae5d4de8048f3ecfd71bab549bcec773459fcdd2dd23d868347", "12");
        Verify("integer-text-reference-changed.numbers", "5f335e1dfea1dd1fdba85bf8984f8a1526c71f64d0293da62764a68834f3e6ee", "18");
        Console.WriteLine("PASS | native Numbers integer-text references | four direct scalar references, native text types/caches, strict saved XLSX and text-operand recalculation; wider coercion remains partial");
    }

    private static void Verify(string file, string expectedHash, string text) {
        string path = Path.Combine(AppContext.BaseDirectory, file);
        Require(Hash(path) == expectedHash, "Native integer-text source changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var input = result.Projection.Sheets[0].Tables.Single().GetCell(4, 2)!;
        Require(input.Kind == IWorkCellKind.Text && input.Value is string value && value == text,
            "Native numeric text lost its scalar type.");
        var table = result.Projection.Sheets[1].Tables.Single();
        Require(table.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete) == 10
            && result.Report.Diagnostics.Any(d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL"),
            "Integer-text admission crossed its qualified boundary.");
        foreach (int row in new[] { 5, 6 })
            Require(!table.GetCell(row, 2)!.FormulaIsComplete && !table.GetCell(row, 3)!.FormulaIsComplete,
                "Unqualified native coercion became complete.");
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        Require(reopened.Sheets.Select(s => s.Name).SequenceEqual(new[] { "Inputs", "Checks" })
            && reopened.Sheets[0].CellAt(4, 2).GetValue<string>() == text, "Saved source type or table identity changed.");
        Require(reopened.Sheets[1].GetFormulaText(4, 2) == "GCD('Inputs'!$B$4,6)"
            && reopened.Sheets[1].GetFormulaText(4, 3) == "LCM('Inputs'!$B$4,6)"
            && reopened.Sheets[1].GetFormulaText(7, 2) == "GCD('Inputs'!$B$4)"
            && reopened.Sheets[1].GetFormulaText(7, 3) == "LCM('Inputs'!$B$4)", "Saved scalar reference changed.");
        Require(reopened.Sheets[1].CellAt(4, 3).GetValue<double>() == (text == "12" ? 12d : 18d),
            "Native formula cache changed.");
        Require(reopened.Calculate() == 10, "Integer-text calculation did not use the shared evaluator.");
        reopened.Sheets[0].CellValue(4, 2, "18");
        Require(reopened.Calculate() == 10 && reopened.Sheets[1].CellAt(4, 2).GetValue<double>() == 6d
            && reopened.Sheets[1].CellAt(4, 3).GetValue<double>() == 18d
            && reopened.Sheets[1].CellAt(7, 2).GetValue<double>() == 18d
            && reopened.Sheets[1].CellAt(7, 3).GetValue<double>() == 18d, "Changed numeric text did not recalculate.");
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Require(!final.ValidateOpenXml().Any() && final.Sheets[0].CellAt(4, 2).GetValue<string>() == "18"
            && final.Sheets[1].CellAt(4, 3).GetValue<double>() == 18d && Hash(path) == expectedHash,
            "Saved edited output or immutable native source changed.");
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
}
