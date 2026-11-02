using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersIntegerFunctionSmoke {
    internal static void Run() {
        string source = Path.Combine(AppContext.BaseDirectory, "integer-functions.numbers");
        const string sourceHash = "99F2F0445B9985FEAC658FA876502C854452026E383F078D41D41260A9CFC0B1";
        string nativeXlsx = Path.Combine(AppContext.BaseDirectory, "integer-functions-native.xlsx");
        const string xlsxHash = "0A6259A8DF8FC6119DC90D8EB814B3E3B4F8F43531A3CC627D15A959A86A904D";
        Require(Hash(source) == sourceHash && Hash(nativeXlsx) == xlsxHash, "Integer-function reference changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable table = result.Projection.Sheets.Single().Tables.Single();
        Require(table.Cells.Count(c => c.Kind == IWorkCellKind.Formula) == 70
            && table.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete) == 39,
            "Integer-function qualification boundary changed.");
        Require(!table.GetCell(10, 6)!.FormulaIsComplete && (double)table.GetCell(10, 6)!.Value! == 0
            && !table.GetCell(14, 7)!.FormulaIsComplete
            && result.Report.Diagnostics.Any(d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL"),
            "Incompatible native coercion/precision caches were made editable.");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Require(reopened.Calculate() == 39 && reopened.Sheets[0].Cell(2, 7).GetValue<double>() == 12
            && reopened.Sheets[0].Cell(2, 8).GetValue<double>() == 360,
            "Qualified integer literals and numeric references did not recalculate after save/reopen.");
        reopened.Sheets[0].Cell(2, 2).SetValue(30d);
        reopened.Calculate(); Require(reopened.Sheets[0].Cell(2, 5).GetValue<double>() == 6
            && reopened.Sheets[0].Cell(2, 6).GetValue<double>() == 180,
            "Edited integer operands did not reach the shared evaluator.");
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Require(final.Sheets[0].Cell(2, 5).GetValue<double>() == 6
            && final.Sheets[0].Cell(2, 6).GetValue<double>() == 180 && !final.ValidateOpenXml().Any(),
            "Saved integer edit or OpenXML validation failed.");
        using var native = ExcelDocument.Load(nativeXlsx);
        Require(native.Calculate() == 68 && native.Sheets[0].Cell(10, 6).GetValue<double>() == 360
            && native.Sheets[0].TryGetCachedFormulaValue(15, 8, out string? overflow) && overflow == "#NUM!",
            "Shared integer evaluation disagrees with the fresh Excel reference.");
        using var bounds = ExcelDocument.Create();
        ExcelSheet guard = bounds.AddSheet("Qualified inputs");
        guard.Cell(1, 4).SetValue(42d); guard.Cell(1, 4).SetFormula("IFERROR(LCM(1E20,OFFSET(A1,-1,0)),99)");
        guard.Cell(2, 4).SetValue(42d); guard.Cell(2, 4).SetFormula("GCD(A1:A100000,B1)");
        guard.Cell(3, 4).SetFormula("IFERROR(SUM(OFFSET(A1,-1,0)),99)");
        guard.Cell(4, 4).SetValue(42d); guard.Cell(4, 4).SetFormula("IFERROR(COUNTIFS(OFFSET(A1,-1,0),UNKNOWN()+1),99)");
        guard.Cell(1, 1).SetValue(1d); guard.Cell(1, 2).SetValue(10d);
        guard.Cell(5, 4).SetFormula("IFNA(XLOOKUP(2,A1:A1,B1:B1,,0,-1),99)");
        Require(bounds.Calculate() == 2 && guard.TryGetCachedFormulaValue(1, 4, out string? unsupported) && unsupported == "42"
            && guard.TryGetCachedFormulaValue(2, 4, out string? overBudget) && overBudget == "42"
            && guard.TryGetCachedFormulaValue(3, 4, out string? referenceError) && referenceError == "99"
            && guard.TryGetCachedFormulaValue(4, 4, out string? criterion) && criterion == "42"
            && guard.TryGetCachedFormulaValue(5, 4, out string? notFound) && notFound == "99",
            "Shared reference qualification or aggregate cell budget failed.");
        Require(Hash(source) == sourceHash && Hash(nativeXlsx) == xlsxHash, "Integer conversion changed native input.");
        Console.WriteLine("PASS | native Numbers integer functions | " + sourceHash
            + " | 19 bounded literal calls, 12 populated numeric ranges, incompatible cache retention, saved/reopened operand edits, native Excel coercion, typed overflow, shared reference qualification and cell budget");
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
