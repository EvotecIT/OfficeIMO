using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersConditionalAggregateSmoke {
    internal static void Run() {
        string source = Fixture("conditional-aggregates.numbers", "A3EA9206083D1062D1D8E190D87350E7A60291AA2600254341BB433DE9007BB1");
        string native = Fixture("conditional-aggregates-excel-native.xlsx", "0D079E885E578CAE87260AEEA05B063C7646A30F349E40CE3D1B8F369A5DEB90");
        string edited = Fixture("conditional-aggregates-excel-edited.xlsx", "9F706AAFD6278037B87CCC340DA3CCE88DA67C72D7098B0F8BBB0DD12A3AD79B");
        using var originalReference = ExcelDocument.Load(native);
        using var editedReference = ExcelDocument.Load(edited);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var table = result.Projection.Sheets.Single().Tables.Single();
        Require(!result.IsVisualFallback && result.Report.IsPartialEditableReconstruction
            && result.Report.Diagnostics.Any(d => d.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED")
            && !result.Report.Diagnostics.Any(d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL")
            && table.Cells.Count(c => c.Kind == IWorkCellKind.Formula) == 20
            && table.Cells.Where(c => c.Kind == IWorkCellKind.Formula).All(c => c.FormulaIsComplete && c.CachedValueIsComplete),
            "Conditional formulas/caches or explicit partial padding boundary changed.");
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        var destination = reopened.Sheets.Single();
        Verify(destination, originalReference.Sheets[0]);
        Require(reopened.Calculate() == 20, "Original conditional formulas were not all recalculated.");
        Verify(destination, originalReference.Sheets[0]);
        destination.CellValue(2, 3, 40d);
        Require(reopened.Calculate() == 20, "Edited conditional formulas were not all recalculated.");
        Verify(destination, editedReference.Sheets[0]);
        using var saved = new MemoryStream(); reopened.Save(saved); saved.Position = 0;
        using var final = ExcelDocument.Load(saved);
        Verify(final.Sheets[0], editedReference.Sheets[0]);
        Require(!final.ValidateOpenXml().Any(), "Saved conditional aggregate workbook is invalid.");
        VerifyEmptyAverageErrors();
        VerifyTypedErrors();
        VerifyUnsupportedCriteria();
        Fixture("conditional-aggregates.numbers", "A3EA9206083D1062D1D8E190D87350E7A60291AA2600254341BB433DE9007BB1");
        Console.WriteLine("PASS | native Numbers conditional aggregates | A3EA9206083D1062D1D8E190D87350E7A60291AA2600254341BB433DE9007BB1"
            + " | twenty saved/reopened formulas, fresh native Excel original/edited results, sixteen empty-average and 397 typed-error results, unsupported criteria guard");
    }

    private static void Verify(ExcelSheet actual, ExcelSheet expected) {
        for (int row = 2; row <= 21; row++) {
            double value = expected.CellAt(row, 7).GetValue<double>();
            Require(actual.GetFormulaText(row, 7) == expected.GetFormulaText(row, 7)
                && Math.Abs(actual.CellAt(row, 7).GetValue<double>() - value) <= 1e-12 * Math.Max(1d, Math.Abs(value)),
                "Conditional formula or value disagrees with native Excel at G" + row);
        }
    }

    private static void VerifyEmptyAverageErrors() {
        string input = Fixture("conditional-average-errors-input.xlsx", "886D1FB2B36ABE335DCCF0F6C01415867D51D70A27B6609B4D9B9C10737C8843");
        string native = Fixture("conditional-average-errors-native.xlsx", "F6698D35704EA12DC7360003A354A072D18DBEA828DC8453195FC558275A4BC7");
        using var reference = ExcelDocument.Load(native);
        using var document = ExcelDocument.Load(input);
        Require(document.Calculate() == 16, "Typed empty-average results or dependent handlers were not calculated.");
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        for (int row = 2; row <= 5; row++) {
            for (int column = 5; column <= 8; column++) {
                Require(reference.Sheets[0].TryGetCachedFormulaValue(row, column, out string? expected)
                    && reopened.Sheets[0].TryGetCachedFormulaValue(row, column, out string? actual)
                    && actual == expected
                    && reopened.Sheets[0].GetFormulaText(row, column) == reference.Sheets[0].GetFormulaText(row, column),
                    "Saved typed average error or handler disagrees with native Excel.");
            }
        }
        reopened.Sheets[0].CellValue(4, 2, 30d);
        Require(reopened.Calculate() == 16 && reopened.Sheets[0].CellAt(4, 5).GetValue<double>() == 30
            && reopened.Sheets[0].CellAt(5, 5).GetValue<double>() == 30
            && !reopened.Sheets[0].CellAt(4, 8).GetValue<bool>()
            && !reopened.ValidateOpenXml().Any(), "Conditional average errors did not recover after an operand edit.");
    }

    private static void VerifyUnsupportedCriteria() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Qualification");
        sheet.CellValue(1, 1, 1d); sheet.CellValue(1, 2, 1d);
        string[] formulas = { "IFERROR(AVERAGEIF(OFFSET(A1,-1,0),UNKNOWN()+1),99)",
            "IFERROR(AVERAGEIFS(B1,OFFSET(A1,-1,0),UNKNOWN()+1),99)",
            "IFERROR(SUMIFS(B1,A1,NA(),A1,UNKNOWN()),99)",
            "IFERROR(COUNTIFS(A1,NA(),A1:A100000,1),99)" };
        for (int index = 0; index < formulas.Length; index++) {
            sheet.CellValue(index + 1, 4, 42d);
            sheet.CellFormula(index + 1, 4, formulas[index]);
        }
        Require(document.Calculate() == 0, "An unsupported conditional criterion was hidden by a reference error.");
        for (int row = 1; row <= formulas.Length; row++)
            Require(sheet.TryGetCachedFormulaValue(row, 4, out string? value) && value == "42",
                "An unsupported conditional criterion replaced its existing cache.");
    }

    private static void VerifyTypedErrors() {
        string input = Fixture("conditional-typed-errors-input.xlsx", "264F8B97C690E4D9541A88E808FD7345222286FADF65FDBB642313B633C264F0");
        string native = Fixture("conditional-typed-errors-native.xlsx", "E42F84E9A886209523604E12BD2EF82BD08704B0D66320E97BCCAE2E05336D80");
        using var reference = ExcelDocument.Load(native);
        using var document = ExcelDocument.Load(input);
        Require(document.Calculate() == 397, "Conditional typed-error cases were not all calculated.");
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        var destination = reopened.Sheets[0];
        for (int row = 2; row <= 99; row++)
            for (int column = 7; column <= 10; column++) VerifyCell(row, column);
        foreach (var cell in new[] { (2, 2), (5, 1), (2, 4), (10, 1), (11, 1) }) VerifyCell(cell.Item1, cell.Item2);
        destination.CellFormula(2, 2, "5");
        Require(reopened.Calculate() == 397 && !reopened.ValidateOpenXml().Any(), "Conditional error recovery was not calculated or saved correctly.");
        for (int row = 2; row <= 7; row++)
            Require(destination.CellAt(row, 7).GetValue<double>() == 5
                && !destination.CellAt(row, 10).GetValue<bool>(), "A selected error did not recover after editing its operand.");

        void VerifyCell(int row, int column) {
            Require(reference.Sheets[0].TryGetCachedFormulaValue(row, column, out string? expected)
                && destination.TryGetCachedFormulaValue(row, column, out string? actual) && actual == expected
                && destination.GetFormulaText(row, column) == reference.Sheets[0].GetFormulaText(row, column),
                "Typed error, criterion identity or dependent handler disagrees with native Excel.");
        }
    }

    private static string Fixture(string name, string hash) {
        string path = Path.Combine(AppContext.BaseDirectory, name);
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash, name + " fixture changed.");
        return path;
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
