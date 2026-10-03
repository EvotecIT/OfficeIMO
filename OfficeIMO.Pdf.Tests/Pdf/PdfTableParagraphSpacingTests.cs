using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableParagraphSpacingTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Fixed_rows_keep_a_fitting_text_line_when_trailing_paragraph_spacing_does_not_fit(string mode) {
        var first = new[] { PdfTextRun.Normal("Visible") };
        var cells = new[] {
            new[] { new PdfTableCell(first, new[] { new PdfTableCellParagraph(first, spacingAfter: 10) }) },
            new[] { new PdfTableCell("Following") }
        };
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        var page = pdf.GetPage(1);
        Assert.Contains("Visible", page.Text);
        Assert.Contains("Following", page.Text);
        var firstGlyph = page.Letters.First(letter => letter.Value == "V");
        var nextGlyph = page.Letters.First(letter => letter.Value == "F");
        Assert.InRange(firstGlyph.StartBaseLine.Y - nextGlyph.StartBaseLine.Y, 15.9, 16.1);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Paragraph_gap_keeps_the_previous_line_visible_but_does_not_pull_the_next_line_into_a_fixed_row(string mode) {
        var first = new[] { PdfTextRun.Normal("Visible") };
        var second = new[] { PdfTextRun.Normal("Outside") };
        var paragraphs = new[] {
            new PdfTableCellParagraph(first, spacingAfter: 8),
            new PdfTableCellParagraph(second, spacingBefore: 6)
        };
        var cells = new[] {
            new[] { new PdfTableCell(first.Concat(second), paragraphs) },
            new[] { new PdfTableCell("Following") }
        };
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        Assert.Contains("Visible", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Outside", pdf.GetPage(1).Text);
        Assert.Contains("Following", pdf.GetPage(1).Text);
    }

    private static byte[] Render(string mode, PdfTableCell[][] cells) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FontSize = 10;
        style.LineHeight = 1.2;
        style.FixedRowHeights = new List<double?> { 16, 16 };
        style.CellPaddingX = 0; style.CellPaddingY = 0;
        style.CellSpacing = 0; style.SpacingBefore = 0; style.SpacingAfter = 0;
        return mode switch {
            "flow" => PdfDocument.Create(options).Table(cells, style: style).ToBytes(),
            "column" => PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)))))).ToBytes(),
            "canvas" => PdfDocument.Create(options).Canvas(canvas => canvas.Table(cells, 24, 24, 192, 32, style)).ToBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(mode))
        };
    }
}
