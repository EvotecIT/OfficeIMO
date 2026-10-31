using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Independent_perimeter_cells_clip_vertical_margins_and_omit_empty_ink_frames(string mode) {
        byte[] Create(bool hidden) => Render(mode, new PdfTableCell("A"), PdfCellVerticalAlign.Top, style => {
            style.ClipTextToCellBounds = true;
            style.PreservePartialCellLines = true;
            style.CellPaddingFromBorderInterior = true;
            style.CellPaddingBottom = global::OfficeIMO.OfficeLength.Points(hidden ? 24D : 0D);
            style.BorderFrame = new PdfTableBorderFrame { Border = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.Black, Width = global::OfficeIMO.OfficeLength.Points(4D) } };
            style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
                [(0, 0)] = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.Black, Width = global::OfficeIMO.OfficeLength.Points(4D), PaintInsideFrame = true }
            };
        });
        using var hidden = PdfPigDocument.Open(Create(true));
        using var partial = PdfPigDocument.Open(Create(false));
        Assert.DoesNotContain("A", hidden.GetPage(1).Text);
        Assert.Contains("A", partial.GetPage(1).Text);
    }

    [Theory]
    [InlineData("flow", 0)]
    [InlineData("column", 0)]
    [InlineData("canvas", 0)]
    [InlineData("flow", -90)]
    [InlineData("column", -90)]
    [InlineData("canvas", -90)]
    [InlineData("flow", 90)]
    [InlineData("column", 90)]
    [InlineData("canvas", 90)]
    public void Inside_frame_borders_clip_at_their_paint_interior_and_retain_partial_text(string mode, int rotation) {
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal("A") }, paragraphs: null, textRotation: rotation);
        byte[] Create(bool hidden) => Render(mode, cell, PdfCellVerticalAlign.Top, style => {
            style.ClipTextToCellBounds = true;
            style.PreservePartialCellLines = true;
            style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
                [(0, 0)] = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.Black, Width = global::OfficeIMO.OfficeLength.Points(4D), PaintInsideFrame = true }
            };
            if (rotation == 0) style.CellPaddingTop = global::OfficeIMO.OfficeLength.Points(hidden ? 20D : 16D);
            else if (rotation < 0) style.CellPaddingRight = global::OfficeIMO.OfficeLength.Points(hidden ? 96D : 92D);
            else style.CellPaddingLeft = global::OfficeIMO.OfficeLength.Points(hidden ? 96D : 92D);
        });
        using var hidden = PdfPigDocument.Open(Create(true));
        using var partial = PdfPigDocument.Open(Create(false));
        Assert.DoesNotContain("A", hidden.GetPage(1).Text);
        Assert.Contains("A", partial.GetPage(1).Text);
    }
}
