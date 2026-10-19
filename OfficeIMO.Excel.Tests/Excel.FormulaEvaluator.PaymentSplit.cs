using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelPaymentSplitTests {
    [Fact]
    public void Payment_splits_match_independent_consumer_caches_after_recalculation() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", "libreoffice-readback.xlsx");
        using var document = ExcelDocument.Load(path);
        var sheet = document.Sheets[0];
        double[] expected = Enumerable.Range(2, 7).Select(row => sheet.CellAt(row, 2).GetValue<double>()).ToArray();
        Assert.Equal(7, document.Calculate());
        for (int index = 0; index < expected.Length; index++)
            Assert.InRange(Math.Abs(expected[index] - sheet.CellAt(index + 2, 2).GetValue<double>()), 0, 1e-10);
    }

    [Theory]
    [InlineData("IPMT(0.1/12,1,36,8000)", -66.66666666666667)]
    [InlineData("IPMT(0.1,3,3,8000)", -292.4471299093655)]
    [InlineData("PPMT(0.1/12,1,24,2000)", -75.6231860083664)]
    [InlineData("IPMT(0,1,12,1200)", 0)]
    [InlineData("PPMT(0,1,12,1200)", -100)]
    [InlineData("IPMT(0.01,1,12,1200,0,1)", 0)]
    public void Payment_splits_match_reference_values_and_survive_save(string formula, double expected) {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Loan");
        sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
        Assert.Equal(1, document.Calculate());
        Assert.InRange(Math.Abs(sheet.CellAt(1, 1).GetValue<double>() - expected), 0, 1e-10);
        using var saved = new MemoryStream(); document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.InRange(Math.Abs(reopened.Sheets[0].CellAt(1, 1).GetValue<double>() - expected), 0, 1e-10);
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData("IPMT(0.01,0,12,1200)", "#NUM!")]
    [InlineData("PPMT(0.01,13,12,1200)", "#NUM!")]
    [InlineData("IPMT(0.01,1,12,1200,0,2)", "#NUM!")]
    [InlineData("IFERROR(IPMT(0.01,0,12,1200),99)", "99")]
    [InlineData("IPMT(#REF!,1,12,1200)", "#REF!")]
    public void Payment_splits_replace_stale_caches_with_typed_errors(string formula, string expected) {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Loan");
        sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
        Assert.Equal(1, document.Calculate());
        Assert.Equal(expected, Assert.Single(document.InspectFormulas().Formulas).CachedValue);
    }

    [Fact]
    public void Loan_schedule_recalculates_dependencies_and_keeps_unsupported_signatures_unevaluated() {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Loan");
        sheet.CellValue(1, 1, 0.01d); sheet.CellValue(2, 1, 1200d);
        sheet.CellFormula(1, 2, "IPMT(A1,2,12,A2,,1)"); sheet.CellFormula(2, 2, "PPMT(A1,2,12,A2,,1)");
        sheet.CellFormula(3, 2, "B1+B2-PMT(A1,12,A2,0,1)");
        Assert.Equal(3, document.Calculate()); Assert.InRange(Math.Abs(sheet.CellAt(3, 2).GetValue<double>()), 0, 1e-12);
        sheet.CellValue(2, 1, 2400d); Assert.Equal(3, document.Calculate());
        Assert.InRange(Math.Abs(sheet.CellAt(3, 2).GetValue<double>()), 0, 1e-12);
        sheet.CellValue(4, 2, 42d); sheet.CellFormula(4, 2, "IPMT(A1:A2,1,12,1200)");
        document.Calculate(); Assert.Equal(42d, sheet.CellAt(4, 2).GetValue<double>());
        Assert.Contains("IPMT", document.InspectFormulas().Capabilities.SupportedFunctions);
        Assert.Contains("PPMT", document.InspectFormulas().Capabilities.SupportedFunctions);
    }
}
