using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Overlapping_viewport_spans_move_as_one_visible_group(string mode) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180, MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0; style.FontSize = 10; style.CellPaddingX = 0; style.CellPaddingY = 0;
        style.CellSpacing = 0; style.SpacingBefore = 0; style.SpacingAfter = 0;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 80 };
        style.FixedRowHeights = new List<double?> { 40, 40, 40 };
        var viewport = new PdfTableCellViewport(50, 160, 50, 80, offsetY: 80);
        var rows = new[] {
            new[] { new PdfTableCell("FirstSpan", rowSpan: 2).WithViewport(viewport), new PdfTableCell("Row1"), new PdfTableCell("Last1") },
            new[] { new PdfTableCell("SecondSpan", rowSpan: 2).WithViewport(viewport), new PdfTableCell("Row2") },
            new[] { new PdfTableCell("Row3"), new PdfTableCell("Last3") }
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(PdfDocument.Create(options).Compose(document => document.Page(page => page.Content(content => {
            content.Item(item => item.Paragraph(paragraph => paragraph.Text("BeforeTable")).Spacer(30));
            if (mode == "flow") content.Item(item => item.Table(rows, style: style));
            else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
        }))).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Row1", pdf.GetPage(1).Text);
        Assert.Contains("Row1", pdf.GetPage(2).Text);
        Assert.Contains("Row2", pdf.GetPage(2).Text);
        Assert.Contains("Row3", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Fractional_fragment_dimensions_remain_contained_without_subtraction_roundoff() {
        var horizontal = new PdfTableCellViewport(15.1 + 15.2, 24, 15.2, 24, offsetX: 15.1);
        var vertical = new PdfTableCellViewport(100, 15.1 + 15.2, 100, 15.2, offsetY: 15.1);
        Assert.Equal(15.1, horizontal.OffsetX);
        Assert.Equal(15.1, vertical.OffsetY);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(30.3, 24, 15.2, 24, offsetX: 15.101));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(100, 30.3, 100, 15.2, offsetY: 15.101));
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Oversized_visible_row_span_is_rejected_before_its_text_leaves_the_page(string mode) {
        ArgumentException error = Assert.Throws<ArgumentException>(() => RenderViewportSpan(mode, 100, 0, false));
        Assert.Contains("viewport", error.Message);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    public void Fitting_visible_row_span_moves_together_including_repeated_headers(string mode, bool repeatHeader) {
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderViewportSpan(mode, 40, 60, repeatHeader));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("BottomFragment", pdf.GetPage(1).Text);
        Assert.Contains("BottomFragment", pdf.GetPage(2).Text);
        Assert.Contains("Row1", pdf.GetPage(2).Text);
        Assert.Contains("Row2", pdf.GetPage(2).Text);
        Assert.All(pdf.GetPage(2).Letters, letter => Assert.True(letter.BoundingBox.Bottom >= 24 - 0.01));
        if (repeatHeader) Assert.Contains("Header", pdf.GetPage(2).Text);
    }

    private static byte[] RenderViewportSpan(string mode, double rowHeight, double prefixHeight, bool repeatHeader) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = repeatHeader ? 1 : 0;
        style.RepeatHeaderRowCount = repeatHeader ? 1 : 0;
        style.FontSize = 10; style.LineHeight = 1.2;
        style.ColumnWidthPoints = new List<double?> { 100, 80 };
        style.FixedRowHeights = repeatHeader ? new List<double?> { 20, rowHeight, rowHeight } : new List<double?> { rowHeight, rowHeight };
        style.CellPaddingX = 0; style.CellPaddingY = 0; style.CellSpacing = 0;
        style.SpacingBefore = 0; style.SpacingAfter = 0;
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };
        var cell = new PdfTableCell("BottomFragment", rowSpan: 2)
            .WithViewport(new PdfTableCellViewport(100, rowHeight * 4, 100, rowHeight * 2, offsetY: rowHeight * 2));
        var rows = new List<PdfTableCell[]>();
        if (repeatHeader) rows.Add(new[] { new PdfTableCell("Header"), new PdfTableCell("Other") });
        rows.Add(new[] { cell, new PdfTableCell("Row1") });
        rows.Add(new[] { new PdfTableCell("Row2") });
        return PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content => {
            if (prefixHeight > 0) content.Item(item => item.Paragraph(paragraph => paragraph.Text("BeforeTable")).Spacer(prefixHeight));
            if (mode == "flow") content.Item(item => item.Table(rows, style: style));
            else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
        }))).ToBytes();
    }
}
