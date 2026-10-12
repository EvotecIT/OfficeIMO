using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersIntegerReferenceSmoke {
    internal static void Run() {
        string path = Path.Combine(AppContext.BaseDirectory, "integer-references.numbers");
        const string expectedHash = "A09E06AF365CFBE210DF7C1F5DC781630061DF61C0E74ED2571A9E0C834B4970";
        Require(Hash(path) == expectedHash, "Native integer reference fixture changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable table = result.Projection.Sheets.Single().Tables.Single();
        Require(table.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete) == 33
            && !table.GetCell(10, 6)!.FormulaIsComplete && (double)table.GetCell(10, 6)!.Value! == 0,
            "Numeric reference or incompatible numeric-text qualification changed.");
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        ExcelSheet sheet = reopened.Sheets.Single();
        Require(sheet.GetFormulaText(2, 7) == "GCD(B2,C2,D2)"
            && sheet.GetFormulaText(2, 8) == "LCM($B$2,$C$2:$D$2,1)"
            && sheet.CellAt(2, 7).GetValue<double>() == 6 && sheet.CellAt(2, 8).GetValue<double>() == 180
            && reopened.Calculate() == 33, "Native reference forms/caches did not survive saved conversion.");
        sheet.CellValue(2, 2, 24d);
        Require(reopened.Calculate() == 33 && sheet.CellAt(2, 5).GetValue<double>() == 12
            && sheet.CellAt(2, 7).GetValue<double>() == 12 && sheet.CellAt(2, 6).GetValue<double>() == 360
            && sheet.CellAt(2, 8).GetValue<double>() == 360, "Saved reference operands did not recalculate.");
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Require(final.Sheets[0].CellAt(2, 7).GetValue<double>() == 12
            && final.Sheets[0].CellAt(2, 8).GetValue<double>() == 360 && !final.ValidateOpenXml().Any()
            && Hash(path) == expectedHash, "Final integer reference save or input immutability failed.");
        Console.WriteLine("PASS | native Numbers integer references | " + expectedHash
            + " | numeric ranges, single/absolute/mixed reference arguments, native caches, saved operand edits and incompatible-cache retention");
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
