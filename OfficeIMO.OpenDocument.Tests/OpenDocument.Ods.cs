using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdsTests {
    [Theory]
    [InlineData("libreoffice-calc-basic.ods")]
    [InlineData("microsoft-excel-basic.ods")]
    public void PreservesAuthoredSpreadsheetFixtureOutsideEditedContent(string fixtureName) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", fixtureName);
        using OdsDocument document = OdsDocument.Open(path);
        var untouched = document.Package.Entries
            .Where(entry => entry.Name != "content.xml" && entry.Name != "META-INF/manifest.xml")
            .ToDictionary(entry => entry.Name, entry => entry.GetOriginalBytes());
        OdsSheet sheet = document.Sheets.First(item => item.UsedRange.HasValue);
        OdsUsedRange used = sheet.UsedRange!.Value;

        sheet.Cell(used.FirstRow, used.FirstColumn).SetString("OfficeIMO");
        byte[] output = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });

        using OdsDocument reopened = OdsDocument.Open(new MemoryStream(output));
        Assert.Equal("OfficeIMO", reopened.GetSheet(sheet.Name)!.GetValue(used.FirstRow, used.FirstColumn).DisplayText);
        foreach (var pair in untouched) Assert.Equal(pair.Value, reopened.Package.GetRequiredEntry(pair.Key).GetOriginalBytes());
    }

    [Fact]
    public void OpensStaticExtremeRepeatFixtureWithoutExpansion() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "extreme-repeats.ods");
        using OdsDocument document = OdsDocument.Open(path);
        OdsSheet sheet = document.Sheets.Single();

        Assert.Equal(2, sheet.RowRuns.Count);
        Assert.Equal(1_048_575, sheet.RowRuns[0].RepeatCount);
        Assert.Equal(16_383, sheet.RowRuns[1].CellRuns[0].RepeatCount);
        Assert.Equal("edge", sheet.GetValue(1_048_575, 16_383).DisplayText);
    }

    [Fact]
    public void EditsExtremeSparseCoordinatesBySplittingRunsWithoutExpansion() {
        using OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Sparse");

        sheet.Cell(1_000_000, 500_000).SetString("edge");

        Assert.Equal(2, sheet.RowRuns.Count);
        Assert.Equal(1_000_000, sheet.RowRuns[0].RepeatCount);
        Assert.Equal(2, sheet.RowRuns[1].CellRuns.Count);
        Assert.Equal(500_000, sheet.RowRuns[1].CellRuns[0].RepeatCount);
        Assert.Equal(new OdsUsedRange(1_000_000, 500_000, 1_000_000, 500_000), sheet.UsedRange);
        Assert.True(document.ToBytes().Length < 10_000);

        using OdsDocument reopened = OdsDocument.Open(new MemoryStream(document.ToBytes()));
        OdsSheet reopenedSheet = reopened.Sheets.Single();
        Assert.Equal("edge", reopenedSheet.GetValue(1_000_000, 500_000).DisplayText);
        Assert.Equal(2, reopenedSheet.RowRuns.Count);
        Assert.Equal(2, reopenedSheet.RowRuns[1].CellRuns.Count);
    }

    [Fact]
    public void WritesTypedValuesFormulaStylesMergeRangesAndValidation() {
        using OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Label");
        sheet.Cell(0, 1).SetDecimal(12.5m);
        sheet.Cell(0, 2).SetBoolean(true);
        sheet.Cell(0, 3).SetDate(new DateTime(2026, 7, 10));
        sheet.Cell(0, 4).SetTime(TimeSpan.FromHours(14.5));
        sheet.Cell(0, 5).SetDuration(TimeSpan.FromDays(2) + TimeSpan.FromMinutes(3));
        sheet.Cell(0, 6).SetPercentage(0.125m);
        sheet.Cell(0, 7).SetCurrency(99.95m, "EUR");
        sheet.Cell(1, 0).SetHyperlink("OfficeIMO", "https://github.com/EvotecIT/OfficeIMO");
        OdsCell formula = sheet.Cell(1, 1);
        formula.Formula = "of:=SUM([.B1:.B1])";
        formula.SetDecimal(12.5m);

        OdsDataStyle numberStyle = document.AddNumberStyle("Amount", 2);
        formula.NumberFormatName = numberStyle.Name;
        formula.Bold = true;
        OdsValidation validation = document.AddValidation("Positive", "cell-content()>=0");
        formula.ValidationName = validation.Name;
        document.AddNamedRange("Amounts", "$Data.$B$1:$B$2");
        sheet.Merge(3, 0, 2, 3).SetString("Merged");
        sheet.Row(3).Height = OdfLength.Centimeters(1);
        sheet.Row(4).Hidden = true;
        sheet.Column(1).Width = OdfLength.Centimeters(3);
        sheet.Column(7).Hidden = true;
        sheet.PrintRanges = "$Data.$A$1:$H$5";

        byte[] bytes = document.ToBytes();
        Assert.True(document.Validate().IsValid);
        using OdsDocument reopened = OdsDocument.Open(new MemoryStream(bytes));
        OdsSheet actual = reopened.Sheets.Single();
        Assert.Equal(12.5m, actual.GetValue(0, 1).AsDecimal());
        Assert.True(actual.GetValue(0, 2).AsBoolean());
        Assert.Equal(new DateTimeOffset(2026, 7, 10, 0, 0, 0, TimeSpan.Zero).Date, actual.GetValue(0, 3).AsDateTimeOffset().Date);
        Assert.Equal(TimeSpan.FromHours(14.5), actual.GetValue(0, 4).AsTimeSpan());
        Assert.Equal(0.125m, actual.GetValue(0, 6).AsDecimal());
        Assert.Equal("EUR", actual.GetValue(0, 7).CurrencyCode);
        Assert.Equal("of:=SUM([.B1:.B1])", actual.Cell(1, 1).Formula);
        Assert.Equal("Amount", actual.Cell(1, 1).NumberFormatName);
        Assert.Equal("Positive", actual.Cell(1, 1).ValidationName);
        Assert.Equal("Merged", actual.GetValue(3, 0).DisplayText);
        Assert.True(actual.Cell(3, 1).IsCovered);
        Assert.Single(reopened.NamedRanges);
        Assert.Single(reopened.Validations);
        Assert.Equal("$Data.$A$1:$H$5", actual.PrintRanges);
    }
}
