using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdsTests {
    [Fact]
    public void StringCellsUseElementContentAsTheCanonicalValue() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("OfficeIMO  value\twith\na line");

        XElement rawCell = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        Assert.Equal("string", (string?)rawCell.Attribute(OdfNamespaces.Office + "value-type"));
        Assert.Null(rawCell.Attribute(OdfNamespaces.Office + "string-value"));

        OdsCellValue reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()))
            .Sheets.Single().GetValue(0, 0);
        Assert.Equal("OfficeIMO  value\twith\na line", reopened.LexicalValue);
        Assert.Equal("OfficeIMO  value\twith\na line", reopened.DisplayText);
    }

    [Fact]
    public void SequentialDenseEdits_Allow_Backfill_And_ContinueAppending() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (var row = 0; row < 50; row++) {
            for (var column = 0; column < 8; column++) {
                sheet.Cell(row, column).SetString($"R{row}C{column}");
            }
        }

        sheet.Cell(25, 3).SetString("Backfilled");
        sheet.Cell(50, 0).SetString("Appended");

        Assert.Equal(new OdsUsedRange(0, 0, 50, 7), sheet.UsedRange);
        Assert.Equal("R49C7", sheet.GetValue(49, 7).DisplayText);
        Assert.Equal("Backfilled", sheet.GetValue(25, 3).DisplayText);
        Assert.Equal("Appended", sheet.GetValue(50, 0).DisplayText);

        OdsSheet reopened = OdsDocument.Load(new MemoryStream(document.ToBytes())).Sheets.Single();
        Assert.Equal("R49C7", reopened.GetValue(49, 7).DisplayText);
        Assert.Equal("Backfilled", reopened.GetValue(25, 3).DisplayText);
        Assert.Equal("Appended", reopened.GetValue(50, 0).DisplayText);
    }

    [Fact]
    public void CellAnnotationsPreserveTextAuthorDateAndIdentity() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(3, 2);
        DateTimeOffset timestamp = new DateTimeOffset(2026, 8, 9, 10, 30, 0, TimeSpan.FromHours(2));

        OdsAnnotation authored = cell.AddAnnotation(
            "Review  this\tvalue\nbefore release", date: timestamp,
            name: "note-17");
        authored.Creator = "Alice";
        authored.Creator = null;
        authored.Creator = "Alice";

        XElement raw = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Office + "annotation").Single();
        Assert.Equal(new[] {
            OdfNamespaces.Dc + "creator",
            OdfNamespaces.Dc + "date",
            OdfNamespaces.Text + "p"
        }, raw.Elements().Select(element => element.Name));
        Assert.True(document.Validate().IsValid);

        Assert.Equal(new OdsUsedRange(3, 2, 3, 2), document.Sheets.Single().UsedRange);
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsAnnotation annotation = Assert.Single(reopened.Sheets.Single().Cell(3, 2).Annotations);
        Assert.Equal("note-17", annotation.Name);
        Assert.Equal("Alice", annotation.Creator);
        Assert.Equal(timestamp, annotation.Date);
        Assert.Equal("Review  this\tvalue\nbefore release", annotation.Text);
    }

    [Fact]
    public void CellAnnotationPrecedesTextContentAndRejectsASecondAnnotation() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("Visible");

        OdsAnnotation annotation = cell.AddAnnotation("Review", "Alice");

        Assert.Equal(annotation.Text, cell.Annotation!.Text);
        Assert.Throws<InvalidOperationException>(() => cell.AddAnnotation("Second"));
        XElement rawCell = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        Assert.Equal(new[] { OdfNamespaces.Office + "annotation", OdfNamespaces.Text + "p" },
            rawCell.Elements().Select(element => element.Name));
    }

    [Fact]
    public void CellTextCaseChangesOnlyDirectDisplayParagraphs() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("first sentence.");
        cell.AddAnnotation("Do NOT change this note", "Alice Example");

        XElement rawCell = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        rawCell.Add(new XElement(OdfNamespaces.Text + "p", "second sentence."));
        rawCell.SetAttributeValue(OdfNamespaces.Office + "string-value", "first sentence.\nsecond sentence.");

        Assert.True(cell.TransformTextCase(OfficeTextCase.SentenceCase));

        Assert.Equal("First sentence.\nSecond sentence.", cell.Text);
        Assert.Equal("First sentence.\nSecond sentence.", (string?)rawCell.Attribute(OdfNamespaces.Office + "string-value"));
        Assert.Equal("Do NOT change this note", cell.Annotation!.Text);
        Assert.Equal("Alice Example", cell.Annotation.Creator);
    }

    [Fact]
    public void CellTextCaseTransformsStoredOnlyStringValueWithoutMaterializingDisplayParagraphs() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("stored only");
        XElement rawCell = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        rawCell.SetAttributeValue(OdfNamespaces.Office + "string-value", "stored only");
        rawCell.Elements(OdfNamespaces.Text + "p").Remove();

        Assert.True(cell.TransformTextCase(OfficeTextCase.Uppercase));

        Assert.Equal("STORED ONLY", (string?)rawCell.Attribute(OdfNamespaces.Office + "string-value"));
        Assert.Empty(rawCell.Elements(OdfNamespaces.Text + "p"));
        OdsCell reopened = OdsDocument.Load(new MemoryStream(document.ToBytes())).Sheets.Single().Cell(0, 0);
        Assert.Equal("STORED ONLY", reopened.Value.LexicalValue);
        Assert.Equal("STORED ONLY", reopened.Value.ToString());
    }

    [Fact]
    public void MergeRejectsMaterializationBeyondTheConfiguredBoundWithoutMutation() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => sheet.Merge(0, 0, 1_000, 1_000));

        Assert.Contains("configured limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(OdsCellValueKind.Empty, sheet.GetValue(0, 0).Kind);
        Assert.Equal(1, sheet.RowCount);
    }

    [Theory]
    [InlineData("libreoffice-calc-basic.ods")]
    [InlineData("microsoft-excel-basic.ods")]
    public void PreservesAuthoredSpreadsheetFixtureOutsideEditedContent(string fixtureName) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", fixtureName);
        OdsDocument document = OdsDocument.Load(path);
        var untouched = document.Package.Entries
            .Where(entry => entry.Name != "content.xml" && entry.Name != "META-INF/manifest.xml")
            .ToDictionary(entry => entry.Name, entry => entry.GetOriginalBytes());
        OdsSheet sheet = document.Sheets.First(item => item.UsedRange.HasValue);
        OdsUsedRange used = sheet.UsedRange!.Value;

        sheet.Cell(used.FirstRow, used.FirstColumn).SetString("OfficeIMO");
        byte[] output = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(output));
        Assert.Equal("OfficeIMO", reopened.GetSheet(sheet.Name)!.GetValue(used.FirstRow, used.FirstColumn).DisplayText);
        foreach (var pair in untouched) Assert.Equal(pair.Value, reopened.Package.GetRequiredEntry(pair.Key).GetOriginalBytes());
    }

    [Fact]
    public void OpensStaticExtremeRepeatFixtureWithoutExpansion() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "extreme-repeats.ods");
        OdsDocument document = OdsDocument.Load(path);
        OdsSheet sheet = document.Sheets.Single();

        Assert.Equal(3, sheet.RowRuns.Count);
        Assert.Equal(1_048_574, sheet.RowRuns[1].RepeatCount);
        Assert.Equal(16_382, sheet.RowRuns[2].CellRuns[1].RepeatCount);
        Assert.Equal("edge", sheet.GetValue(1_048_575, 16_383).DisplayText);
    }

    [Fact]
    public void EditsExtremeSparseCoordinatesBySplittingRunsWithoutExpansion() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Sparse");

        sheet.Cell(1_000_000, 500_000).SetString("edge");

        Assert.Equal(3, sheet.RowRuns.Count);
        Assert.Equal(999_999, sheet.RowRuns[1].RepeatCount);
        Assert.Equal(3, sheet.RowRuns[2].CellRuns.Count);
        Assert.Equal(499_999, sheet.RowRuns[2].CellRuns[1].RepeatCount);
        Assert.Equal(new OdsUsedRange(1_000_000, 500_000, 1_000_000, 500_000), sheet.UsedRange);
        Assert.True(document.ToBytes().Length < 10_000);

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsSheet reopenedSheet = reopened.Sheets.Single();
        Assert.Equal("edge", reopenedSheet.GetValue(1_000_000, 500_000).DisplayText);
        Assert.Equal(3, reopenedSheet.RowRuns.Count);
        Assert.Equal(3, reopenedSheet.RowRuns[2].CellRuns.Count);
    }

    [Fact]
    public void WritesTypedValuesFormulaStylesMergeRangesAndValidation() {
        OdsDocument document = OdsDocument.Create();
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
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(bytes));
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
