using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_AutoHeightSingleParagraphCellPreservesLargerRichSpan(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var table = document.AddTable(1, 1);
        table.Width = 4200;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new System.Collections.Generic.List<int> { 4200 };
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        var paragraph = table.Rows[0].Cells[0].AddParagraph("Prefix ", removeExistingParagraphs: true);
        AddLargerRichTail(paragraph);
        var options = CreateWideTextOptions(useShaper);

        var snapshot = document.CreateVisualSnapshot(options);
        var image = document.ExportImage(OfficeImageExportFormat.Svg, options);

        AssertLargerRichTail(snapshot);
        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "R", 20);
        Assert.DoesNotContain(image.Diagnostics, item => item.Code == "unsupported-word-table-cell-text-overflow");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_VmlFitShapePreservesLargerRichSpan(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var textBox = document.AddParagraph().AddTextBoxVml("Prefix ");
        var prefix = textBox.Paragraphs[0];
        // The legacy constructor leaves xml:space unspecified. Author the separator
        // through Text so the fixture explicitly preserves the intended run boundary.
        prefix.Text = "Prefix ";
        AddLargerRichTail(prefix);
        V.Shape shape = document.BodyRoot.Descendants<V.Shape>().Last(item => item.Descendants<V.TextBox>().Any());
        shape.Style = "width:210pt;mso-wrap-style:square;mso-fit-shape-to-text:t";
        var options = CreateWideTextOptions(useShaper);
        var authoredTail = Assert.Single(textBox.Paragraphs, paragraph => paragraph.Text.Contains("R20"));
        Assert.Equal(ManagedTextShapingTestAssets.FamilyName, authoredTail.FontFamily);

        var snapshot = document.CreateVisualSnapshot(options);
        var image = document.ExportImage(OfficeImageExportFormat.Svg, options);

        AssertLargerRichTail(snapshot);
        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "R", 20);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_AutoHeightCellAfterNestedTablePreservesLargerRichSpan(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var table = document.AddTable(1, 1);
        table.Width = 4200;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new System.Collections.Generic.List<int> { 4200 };
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        var cell = table.Rows[0].Cells[0];
        var nested = cell.AddTable(1, 1, removePrecedingParagraph: true);
        nested.Width = 3600;
        nested.WidthType = WordTableWidthUnit.Dxa;
        nested.ColumnWidth = new System.Collections.Generic.List<int> { 3600 };
        nested.ColumnWidthType = WordTableWidthUnit.Dxa;
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Nested context";
        AddLargerRichTail(cell.AddParagraph("Prefix "));
        var options = CreateWideTextOptions(useShaper);

        var snapshot = document.CreateVisualSnapshot(options);
        var image = document.ExportImage(OfficeImageExportFormat.Svg, options);

        AssertLargerRichTail(snapshot);
        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "R", 20);
        Assert.Contains("Nested context", Encoding.UTF8.GetString(image.Bytes));
        Assert.DoesNotContain(image.Diagnostics, item => item.Code == "unsupported-word-nested-table-overflow" ||
            item.Code == "unsupported-word-table-cell-text-overflow");
    }

    [Fact]
    public void WordDocument_AutoHeightCellAfterInlineImagePreservesLargerRichSpan() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var table = document.AddTable(1, 1);
        table.Width = 4200;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new System.Collections.Generic.List<int> { 4200 };
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        var cell = table.Rows[0].Cells[0];
        using var imageStream = new MemoryStream(CreateSolidPng(24, 24, OfficeColor.Red));
        cell.AddParagraph(removeExistingParagraphs: true)
            .AddImage(imageStream, "rich-cell.png", 24, 24, description: "Inline context");
        AddLargerRichTail(cell.AddParagraph("Prefix "));
        var options = CreateWideTextOptions(useShaper: false);

        var snapshot = document.CreateVisualSnapshot(options);
        var image = document.ExportImage(OfficeImageExportFormat.Svg, options);

        AssertLargerRichTail(snapshot);
        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "R", 20);
        Assert.Equal("Inline context", Assert.Single(snapshot.Drawing.Images).AlternativeText);
        Assert.DoesNotContain(image.Diagnostics, item => item.Code == "unsupported-word-table-image-overflow" ||
            item.Code == "unsupported-word-table-cell-text-overflow");
    }

    private static void AddLargerRichTail(WordParagraph paragraph) {
        paragraph.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        paragraph.FontSizePoints = 11D;
        var tail = paragraph.AddText(string.Join(" ", Enumerable.Range(1, 20)
            .Select(index => "R" + index.ToString("00", CultureInfo.InvariantCulture))));
        tail.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        tail.FontSizePoints = 22D;
        tail.Bold = true;
    }

    private static void AssertLargerRichTail(WordDocumentVisualSnapshot snapshot) {
        var rich = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
        var prefix = Assert.Single(rich.Runs, run => run.FontSize == 11D);
        Assert.Equal("Prefix ", prefix.Text);
        Assert.Equal(ManagedTextShapingTestAssets.FamilyName, prefix.FontFamily);
        var tail = Assert.Single(rich.Runs, run => run.FontSize == 22D && run.Bold);
        Assert.Equal(ManagedTextShapingTestAssets.FamilyName, tail.FontFamily);
        Assert.Contains("R20", tail.Text);
    }
}
