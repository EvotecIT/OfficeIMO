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
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(240), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(48), CompressContentStreams = false };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ClipTextToCellBounds = containText;
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100) };
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(24) };
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0); style.CellSpacing = global::OfficeIMO.OfficeLength.Points(0);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0); style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0);
        var runs = new[] { new PdfTextRun("A") };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
            leftIndent: -30, lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(6))) });
        var cells = new[] { new[] { cell } };
        var document = PdfDocument.Create(options);
        style = style.Clone();
        byte[] bytes = frame switch {
            "flow" => document.Table(cells, style: style).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "column" => document.Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)))))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "canvas" => document.Canvas(canvas => canvas.Table(cells, 24, 24, 100, 24, style)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
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
