using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, false, 2)]
    [InlineData(true, false, 2)]
    [InlineData(false, true, 2)]
    [InlineData(true, true, 2)]
    [InlineData(false, false, 3)]
    [InlineData(true, false, 3)]
    [InlineData(false, true, 3)]
    [InlineData(true, true, 3)]
    public void Table_RequiredRowHeightMovesIntactWhenOnlyItsTextFitsRemainingSpace(bool inColumn, bool fixedHeight, int lineCount) {
        using var pdf = PdfPigDocument.Open(CreateTableWithRequiredRowHeight(inColumn, fixedHeight, lineCount));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BeforeTable", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Unique", pdf.GetPage(1).Text);
        for (int line = 1; line <= lineCount; line++) Assert.Contains("Unique" + line.ToString("D2"), pdf.GetPage(2).Text);
    }

    private static byte[] CreateTableWithRequiredRowHeight(bool inColumn, bool fixedHeight, int lineCount) {
        var options = new PdfOptions { PageWidth = 360, PageHeight = 210, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0; style.RepeatHeaderRowCount = 0;
        style.MinimumBodyRowsOnFirstPage = 0; style.MinimumBodyRowsOnLastPage = 0;
        style.CellPaddingY = 2; style.SpacingBefore = 0; style.SpacingAfter = 0; style.CellSpacing = 0;
        if (fixedHeight) style.FixedRowHeights = new List<double?> { null, 100 };
        else style.RowMinHeights = new List<double?> { null, 100 };
        var rows = new[] { new[] { "Filler" }, new[] { string.Join("\n", Enumerable.Range(1, lineCount).Select(line => "Unique" + line.ToString("D2"))) } };
        var document = PdfDocument.Create(options);
        if (inColumn) document.Compose(compose => compose.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Paragraph(p => p.Text("BeforeTable")).Spacer(55).Table(rows, style: style))))));
        else document.Paragraph(p => p.Text("BeforeTable")).Spacer(55).Table(rows, style: style);
        return document.ToBytes();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Table_RowHeightRequirementCannotOverflowItsContinuationFrame(bool inColumn) {
        var options = new PdfOptions { PageWidth = 260, PageHeight = 180, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0; style.RepeatHeaderRowCount = 0; style.MinRowHeight = 141;
        var document = PdfDocument.Create(options);
        var rows = new[] { new[] { "Required row height" } };
        if (inColumn) document.Compose(compose => compose.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        else document.Table(rows, style: style);
        ArgumentException error = Assert.Throws<ArgumentException>(() => document.ToBytes());
        Assert.Contains("row height requirement", error.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Table_PartiallyRenderedRowDoesNotRestartOnItsContinuationPage(bool inColumn) {
        var options = new PdfOptions { PageWidth = 360, PageHeight = 210, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1; style.RepeatHeaderRowCount = 1;
        style.CellPaddingY = 2;
        var rows = new[] { new[] { "Heading" }, new[] { string.Join("\n", Enumerable.Range(1, 11).Select(line => "Unique" + line.ToString("D2"))) } };
        var document = PdfDocument.Create(options);
        if (inColumn) document.Compose(compose => compose.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Spacer(50).Table(rows, style: style))))));
        else document.Spacer(50).Table(rows, style: style);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Unique01", pdf.GetPage(1).Text);
        Assert.Contains("Unique11", pdf.GetPage(2).Text);
        string text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        for (int line = 1; line <= 11; line++) Assert.Equal(1, Regex.Matches(text, "Unique" + line.ToString("D2")).Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Table_ContinuationHeadersMeasureRemainingContentAfterFirstSegmentImage(bool inColumn) {
        byte[] bytes = CreateTableWithFirstSegmentImage(inColumn);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages >= 3);
        Assert.Contains("Heading", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Heading", pdf.GetPage(2).Text);
        Assert.Equal(1, pdf.GetPage(2).NumberOfImages);
        for (int pageNumber = 3; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            Assert.Contains("Heading", pdf.GetPage(pageNumber).Text);
            Assert.Equal(0, pdf.GetPage(pageNumber).NumberOfImages);
        }
        foreach (var page in pdf.GetPages()) {
            Assert.All(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)),
                letter => Assert.InRange(letter.StartBaseLine.Y, 18, 160));
        }
        string text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        for (int line = 1; line <= 45; line++) Assert.Contains("Line" + line.ToString("D2"), text);
    }

    private static byte[] CreateTableWithFirstSegmentImage(bool inColumn) {
        var options = new PdfOptions { PageWidth = 260, PageHeight = 180, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1; style.RepeatHeaderRowCount = 1;
        style.CellPaddingX = 2; style.CellPaddingY = 2;
        style.ColumnWidthPoints = new List<double?> { 60, null };
        var cells = new[] {
            new[] { new PdfTableCell("Heading"), new PdfTableCell("Description") },
            new[] { PdfTableCell.WithImages("", new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 30, 130) }),
                new PdfTableCell(string.Join("\n", Enumerable.Range(1, 45).Select(line => "Line" + line.ToString("D2")))) }
        };
        var document = PdfDocument.Create(options);
        if (inColumn) document.Compose(compose => compose.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style))))));
        else document.Table(cells, style: style);
        return document.ToBytes();
    }
}
