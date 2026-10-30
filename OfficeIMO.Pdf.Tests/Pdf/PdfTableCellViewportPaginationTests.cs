using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Overlapping_viewport_spans_move_as_one_visible_group(string mode) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(240), PageHeight = global::OfficeIMO.OfficeLength.Points(180), MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24) };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0; style.FontSize = global::OfficeIMO.OfficeLength.Points(10); style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.CellSpacing = global::OfficeIMO.OfficeLength.Points(0); style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0); style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(80) };
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(40), global::OfficeIMO.OfficeLength.Points(40) };
        var viewport = new PdfTableCellViewport(50, 160, 50, 80, offsetY: 80);
        var rows = new[] {
            new[] { new PdfTableCell("FirstSpan", rowSpan: 2).WithViewport(viewport), new PdfTableCell("Row1"), new PdfTableCell("Last1") },
            new[] { new PdfTableCell("SecondSpan", rowSpan: 2).WithViewport(viewport), new PdfTableCell("Row2") },
            new[] { new PdfTableCell("Row3"), new PdfTableCell("Last3") }
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(PdfDocument.Create(options).Compose(document => document.Page(page => page.Content(content => {
            content.Item(item => item.Paragraph(paragraph => paragraph.Text("BeforeTable")).Spacer(global::OfficeIMO.OfficeLength.Points(30)));
            if (mode == "flow") content.Item(item => item.Table(rows, style: style));
            else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
        }))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(240), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = repeatHeader ? 1 : 0;
        style.RepeatHeaderRowCount = repeatHeader ? 1 : 0;
        style.FontSize = global::OfficeIMO.OfficeLength.Points(10); style.LineHeight = 1.2;
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(80) };
        style.FixedRowHeights = (repeatHeader ? new List<double?> { 20, rowHeight, rowHeight } : new List<double?> { rowHeight, rowHeight }) is { } lengths5090 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths5090, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0); style.CellSpacing = global::OfficeIMO.OfficeLength.Points(0);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0); style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0);
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };
        var cell = new PdfTableCell("BottomFragment", rowSpan: 2)
            .WithViewport(new PdfTableCellViewport(100, rowHeight * 4, 100, rowHeight * 2, offsetY: rowHeight * 2));
        var rows = new List<PdfTableCell[]>();
        if (repeatHeader) rows.Add(new[] { new PdfTableCell("Header"), new PdfTableCell("Other") });
        rows.Add(new[] { cell, new PdfTableCell("Row1") });
        rows.Add(new[] { new PdfTableCell("Row2") });
        return PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content => {
            if (prefixHeight > 0) content.Item(item => item.Paragraph(paragraph => paragraph.Text("BeforeTable")).Spacer(global::OfficeIMO.OfficeLength.Points(prefixHeight)));
            if (mode == "flow") content.Item(item => item.Table(rows, style: style));
            else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
        }))).ToBytes();
    }
}
