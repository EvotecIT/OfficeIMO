using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableTextClipBoundsTests {
    [Theory]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    public void Cell_clip_policy_survives_style_copy_and_retains_authored_glyph_positions(string frame, bool containText) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 48, CompressContentStreams = false };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ClipTextToCellBounds = containText;
        style.ColumnWidthPoints = new List<double?> { 100 };
        style.FixedRowHeights = new List<double?> { 24 };
        style.CellPaddingX = 0; style.CellPaddingY = 0; style.CellSpacing = 0;
        style.SpacingBefore = 0; style.SpacingAfter = 0;
        var runs = new[] { new PdfTextRun("A") };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
            leftIndent: -30, lineSpacing: PdfLineSpacing.Exactly(6)) });
        var cells = new[] { new[] { cell } };
        var document = PdfDocument.Create(options);
        style = style.Clone();
        byte[] bytes = frame switch {
            "flow" => document.Table(cells, style: style).ToBytes(),
            "column" => document.Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)))))).ToBytes(),
            "canvas" => document.Canvas(canvas => canvas.Table(cells, 24, 24, 100, 24, style)).ToBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(frame))
        };
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        string clip = containText ? "24 132 100 24 re W n" :
            frame == "canvas" ? "22 130 104 28 re W n" : "-8 130 134 28 re W n";
        Assert.Contains(clip, content);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        Assert.Equal("A", letter.Value);
        Assert.Equal(-6D, letter.StartBaseLine.X, 3);
    }
}
