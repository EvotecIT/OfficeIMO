using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("columns")]
    [InlineData("canvas")]
    public void TablePairedBorders_OneSidedBoundaryReservesNeighbourText(string surface) {
        var style = MixedPairedBorderStyle();
        style.CellBorders![(0, 0)] = MixedPairedBorder(right: true);
        byte[] bytes = RenderMixedPairedTable(surface, style);
        var lines = PairedBorderLines(bytes);
        using var pdf = PdfPigDocument.Open(bytes);
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "gypB");
        Assert.Equal(2, lines.Length);
        Assert.True(word.BoundingBox.Left >= lines.Max(line => line.X1) + 1,
            "The neighbour must reserve the incoming stroke even without its own border.");
        Assert.Null(style.CellPaddings);
        Assert.Single(style.CellBorders);
        Assert.Null(style.CellBorders[(0, 0)].RightBorder);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TablePairedBorders_DisabledOrNarrowerNeighbourStillReservesIncomingPaint(bool narrower) {
        var style = MixedPairedBorderStyle();
        style.CellBorders![(0, 0)] = MixedPairedBorder(right: true);
        style.CellBorders[(0, 1)] = new PdfCellBorder {
            Color = null, Top = false, Right = false, Bottom = false, Left = narrower,
            LeftBorder = new PdfCellBorderSide { Color = PdfColor.FromRgb(0, 0, 255), Width = .25 }
        };
        using var pdf = PdfPigDocument.Open(RenderMixedPairedTable("flow", style));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "gypB");
        Assert.True(word.BoundingBox.Left >= 133);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    public void TablePairedBorders_RowSpanReservesBothNeighboursWithCellSpacing(double spacing) {
        var style = MixedPairedBorderStyle();
        style.CellSpacing = spacing;
        style.CellBorders![(0, 0)] = MixedPairedBorder(right: true);
        PdfTableCell[][] rows = {
            new[] { new PdfTableCell("gypA", rowSpan: 2), new PdfTableCell("gypB") },
            new[] { new PdfTableCell("gypD") }
        };
        byte[] bytes = RenderMixedPairedTable("flow", style, rows);
        var lines = PairedBorderLines(bytes);
        using var pdf = PdfPigDocument.Open(bytes);
        foreach (string marker in new[] { "gypB", "gypD" }) {
            var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == marker);
            Assert.True(word.BoundingBox.Left >= lines.Max(line => line.X1) + 1);
        }
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("columns")]
    [InlineData("canvas")]
    public void TablePairedBorders_LaterRowFillKeepsBothVisibleTracks(string surface) {
        var style = MixedPairedBorderStyle();
        style.CellBorders![(0, 0)] = MixedPairedBorder(bottom: true);
        style.CellFills = new Dictionary<(int, int), PdfColor> { [(1, 0)] = PdfColor.FromRgb(255, 230, 160) };
        byte[] bytes = RenderMixedPairedTable(surface, style);
        var lines = PairedBorderLines(bytes);
        Assert.Equal(2, lines.Length);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        foreach (var line in lines) {
            int x = (int)Math.Round((line.X1 + line.X2) / 2);
            int y = (int)Math.Round(220 - line.Y1);
            Assert.Equal(OfficeColor.Red, raster.GetPixel(x, y));
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2.5)]
    public void TablePairedBorders_SmallRadiusKeepsInnerTrackAlignedAndUnclipped(double radius) {
        var style = MixedPairedBorderStyle();
        style.CornerRadius = radius;
        style.PreferredWidth = 300;
        style.ColumnWidthPoints = new List<double?> { 100, 100, 100 };
        style.FixedRowHeights = new List<double?> { 20 };
        for (int column = 0; column < 3; column++) style.CellBorders![(0, column)] = MixedPairedBorder(top: true);
        byte[] bytes = RenderMixedPairedTable("flow", style,
            new[] { new[] { new PdfTableCell("gypA"), new PdfTableCell("gypB"), new PdfTableCell("gypC") } });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        for (int y = 27; y < 40; y++) {
            Assert.Equal(raster.GetPixel(180, y), raster.GetPixel(80, y));
            Assert.Equal(raster.GetPixel(180, y), raster.GetPixel(280, y));
        }
        Assert.Equal(OfficeColor.Red, raster.GetPixel(80, 34));
    }

    private static PdfCellBorder MixedPairedBorder(bool top = false, bool right = false, bool bottom = false) => new() {
        Color = PdfColor.FromRgb(255, 0, 0), Width = 2, LineStyle = PdfCellBorderLineStyle.TwoLine,
        Top = top, Right = right, Bottom = bottom, Left = false
    };

    private static PdfTableStyle MixedPairedBorderStyle() => new() {
        HeaderRowCount = 0, BorderColor = null, HeaderFill = null, RowStripeFill = null,
        CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 1, PreferredWidth = 200,
        ColumnWidthPoints = new List<double?> { 100, 100 }, FixedRowHeights = new List<double?> { 20, 20 },
        CellBorders = new Dictionary<(int, int), PdfCellBorder>()
    };

    private static byte[] RenderMixedPairedTable(string surface, PdfTableStyle style, PdfTableCell[][]? rows = null) {
        rows ??= new[] {
            new[] { new PdfTableCell("gypA"), new PdfTableCell("gypB") },
            new[] { new PdfTableCell("gypC"), new PdfTableCell("gypD") }
        };
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 520, PageHeight = 220, MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30,
            CompressContentStreams = false
        });
        document.Compose(root => root.Page(page => page.Content(content => {
            switch (surface) {
                case "flow": content.Table(rows, style: style); break;
                case "row": content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))); break;
                case "columns": content.Columns(inner => inner.Table(rows, style: style), new PdfMultiColumnOptions { ColumnCount = 2, BalanceLastPage = false }); break;
                case "canvas": content.Canvas(canvas => canvas.Table(rows, 30, 30, 200, 40, style)); break;
                default: throw new ArgumentException("Unknown surface", nameof(surface));
            }
        })));
        return document.ToBytes();
    }
}
