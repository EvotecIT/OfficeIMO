using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRoundedCornerTests {
    private static List<PdfTableCell[]> Callout() => new() {
        new[] {
            new PdfTableCell(new[] {
                new PdfTextRun("Callout title", bold: true, color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.1, 0.1, 0.1), fontSize: 11),
                new PdfTextRun("\n", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.1, 0.1, 0.1), fontSize: 5),
                new PdfTextRun("Body with ", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun("bold", bold: true, color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun(" and a ", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun("link", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.1, 0.4, 0.7), fontSize: 10, linkUri: "https://example.com"),
                new PdfTextRun(".", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2, 0.2, 0.2), fontSize: 10),
            }),
        },
    };

    private static PdfTableStyle CalloutStyle(double radius) => new() {
        HeaderRowCount = 0,
        BorderColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.85, 0.85, 0.85),
        BorderWidth = global::OfficeIMO.OfficeLength.Points(1),
        RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0),
        CornerRadius = global::OfficeIMO.OfficeLength.Points(radius),
        CellFills = new Dictionary<(int, int), global::OfficeIMO.OfficeColor> { [(0, 0)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.95, 0.97, 0.99) },
        CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder { LeftBorder = new PdfCellBorderSide { Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.13, 0.47, 0.71), Width = global::OfficeIMO.OfficeLength.Points(4) } },
        },
    };

    [Fact]
    public void TableCornerRadius_RoundsOuterBox_ClipsCellBorder_AndKeepsInlineText() {
        byte[] rounded = PdfDocument.Create(d => d.Content(c => c.Table(Callout(), PdfAlign.Left, CalloutStyle(6)))).ToBytes();
        byte[] square = PdfDocument.Create(d => d.Content(c => c.Table(Callout(), PdfAlign.Left, CalloutStyle(0)))).ToBytes();

        string roundedRaw = PdfTestContentText.ReadDecoded(rounded);
        string squareRaw = PdfTestContentText.ReadDecoded(square);

        // Rounding changes the drawing: the square box strokes a plain rectangle (re), the rounded box
        // emits bezier corners (c) for its outer fill and border and clips the accent stripe to them.
        Assert.Contains(" c", roundedRaw);        // cubic-bezier corner operator, present only when rounded
        Assert.DoesNotContain(" c", squareRaw);   // a square box draws only rectangles, never a bezier
        Assert.NotEqual(squareRaw, roundedRaw);
        Assert.Contains("0.8 0.8 0.8 RG", roundedRaw, StringComparison.Ordinal); // shared fallback sides
        Assert.Contains("0.13 0.47 0.71 RG", roundedRaw, StringComparison.Ordinal); // explicit left side

        // Inline formatting and the link survive the rounding.
        string text = PdfReadDocument.Open(rounded).ExtractText();
        Assert.Contains("Callout title", text, StringComparison.Ordinal);
        Assert.Contains("bold", text, StringComparison.Ordinal);
        Assert.Contains("link", text, StringComparison.Ordinal);
        Assert.Contains("https://example.com", roundedRaw, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelCornerRadius_RoundsBox_AndKeepsText() {
        byte[] bytes = PdfDocument.Create()
            .Container(content => {
                content.H2("Panel title");
                content.Paragraph(p => p.Text("Panel body"));
            }, new PdfPanelStyle {
                Background = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.95, 0.97, 0.99),
                BorderColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2, 0.3, 0.4),
                BorderWidth = 1,
                CornerRadius = 8,
                PaddingX = 12,
                PaddingY = 10,
            })
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains(" c", raw);   // rounded corners emit bezier operators
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Panel title", text, StringComparison.Ordinal);
        Assert.Contains("Panel body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTableCornerRadius_RoundsTheFixedPositionOuterBox() {
        var style = CalloutStyle(8);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(220),
                PageHeight = global::OfficeIMO.OfficeLength.Points(160),
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(Callout(), 24, 24, 160, 64, style))
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.Contains("Callout title", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void DeferredTableCornerRadius_UsesLogicalEdgesRegardlessOfBatchSize() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CornerRadius = global::OfficeIMO.OfficeLength.Points(7);

        byte[] oneRowBatches = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .TableDeferred(CreateRows, batchSize: 1, style: style)
            .ToBytes();
        byte[] threeRowBatches = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .TableDeferred(CreateRows, batchSize: 3, style: style)
            .ToBytes();

        Assert.Equal(
            CountOccurrences(PdfTestContentText.ReadDecoded(oneRowBatches), " c"),
            CountOccurrences(PdfTestContentText.ReadDecoded(threeRowBatches), " c"));

        static IEnumerable<string[]> CreateRows() {
            for (int row = 0; row < 7; row++) yield return new[] { "Row " + row, "Value " + row };
        }
    }

    [Fact]
    public void RoundedCellBorder_PreservesDashDoubleLineAndDiagonals() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.CornerRadius = global::OfficeIMO.OfficeLength.Points(8);
        style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = global::OfficeIMO.OfficeColor.FromRgb(68, 85, 102),
                Width = global::OfficeIMO.OfficeLength.Points(1),
                DashStyle = OfficeStrokeDashStyle.Dash,
                LineStyle = PdfCellBorderLineStyle.TwoLine,
                DiagonalUp = true,
                DiagonalDown = true
            }
        };

        string raw = PdfTestContentText.ReadDecoded(PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Table(new[] { new[] { "Border contract" } }, style: style)
            .ToBytes());

        Assert.Contains("[3 1.5] 0 d", raw, StringComparison.Ordinal);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.True(CountOccurrences(raw, " S") >= 12, "Expected two strokes for every side and both diagonals.");
    }

    [Fact]
    public void CornerRadius_DoesNotRoundInteriorCellCorners() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.FooterFill = null;
        style.RowStripeFill = null;
        style.CornerRadius = global::OfficeIMO.OfficeLength.Points(8);
        style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.FromRgb(68, 85, 102), Width = global::OfficeIMO.OfficeLength.Points(1) }
        };

        string raw = PdfTestContentText.ReadDecoded(PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Table(new[] { new[] { "Left", "Middle", "Right" } }, style: style)
            .ToBytes());
        style.CellBorders = null;
        string withoutCellBorder = PdfTestContentText.ReadDecoded(PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Table(new[] { new[] { "Left", "Middle", "Right" } }, style: style)
            .ToBytes());

        Assert.Equal(CountOccurrences(withoutCellBorder, " c"), CountOccurrences(raw, " c"));
    }

    [Fact]
    public void RoundedRowSpannedCellFill_ReachesItsLogicalBottomEdge() {
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.Merge(string.Empty, rowSpan: 2), PdfTableCell.TextCell(string.Empty) },
            new[] { PdfTableCell.TextCell(string.Empty) }
        };
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.CornerRadius = global::OfficeIMO.OfficeLength.Points(10);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(60) };
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30) };
        style.CellFills = new Dictionary<(int, int), global::OfficeIMO.OfficeColor> { [(0, 0)] = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0) };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(160),
                PageHeight = global::OfficeIMO.OfficeLength.Points(120),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
            })
            .Table(rows, style: style)
            .ToBytes();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(40, 65));
    }

    [Fact]
    public void OversizedTableCornerRadius_UsesOneBoundaryRadiusForPerimeterFillAndCellBorder() {
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.Merge(string.Empty, rowSpan: 2), PdfTableCell.TextCell(string.Empty) },
            new[] { PdfTableCell.TextCell(string.Empty) }
        };

        byte[] expectedFlow = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Table(rows, style: CreateStyle(10))
            .ToBytes();
        byte[] oversizedFlow = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .Table(rows, style: CreateStyle(100))
            .ToBytes();
        byte[] expectedCanvas = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(180), PageHeight = global::OfficeIMO.OfficeLength.Points(120), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10), MarginTop = global::OfficeIMO.OfficeLength.Points(10), MarginBottom = global::OfficeIMO.OfficeLength.Points(10), CompressContentStreams = false })
            .Canvas(canvas => canvas.Table(rows, 20, 20, 120, 60, CreateStyle(10)))
            .ToBytes();
        byte[] oversizedCanvas = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(180), PageHeight = global::OfficeIMO.OfficeLength.Points(120), MarginLeft = global::OfficeIMO.OfficeLength.Points(10), MarginRight = global::OfficeIMO.OfficeLength.Points(10), MarginTop = global::OfficeIMO.OfficeLength.Points(10), MarginBottom = global::OfficeIMO.OfficeLength.Points(10), CompressContentStreams = false })
            .Canvas(canvas => canvas.Table(rows, 20, 20, 120, 60, CreateStyle(100)))
            .ToBytes();

        Assert.Equal(expectedFlow, oversizedFlow);
        Assert.Equal(expectedCanvas, oversizedCanvas);

        static PdfTableStyle CreateStyle(double radius) {
            PdfTableStyle style = PdfTableStyles.Minimal();
            style.HeaderRowCount = 0;
            style.BorderColor = global::OfficeIMO.OfficeColor.FromRgb(68, 85, 102);
            style.BorderWidth = global::OfficeIMO.OfficeLength.Points(1);
            style.CornerRadius = global::OfficeIMO.OfficeLength.Points(radius);
            style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
            style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
            style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(60) };
            style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(40) };
            style.CellFills = new Dictionary<(int, int), global::OfficeIMO.OfficeColor> { [(0, 0)] = global::OfficeIMO.OfficeColor.FromRgb(220, 235, 250) };
            style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
                [(0, 0)] = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.FromRgb(24, 119, 181), Width = global::OfficeIMO.OfficeLength.Points(3) }
            };
            return style;
        }
    }

    [Fact]
    public void CornerRadius_ClonesAndRejectsNegative() {
        PdfTableStyle table = new() { CornerRadius = global::OfficeIMO.OfficeLength.Points(6) };
        Assert.Equal(6, table.Clone().CornerRadius.ToPoints());
        Assert.Throws<ArgumentException>(() => new PdfTableStyle { CornerRadius = global::OfficeIMO.OfficeLength.Points(-1) });

        PdfPanelStyle panel = new() { CornerRadius = 8 };
        Assert.Equal(8, panel.Clone().CornerRadius);
        Assert.Throws<ArgumentException>(() => new PdfPanelStyle { CornerRadius = -1 });
        Assert.Throws<ArgumentException>(() => new PdfPanelStyle { CornerRadius = double.NaN });
        Assert.Throws<ArgumentException>(() => new PdfPanelStyle { CornerRadius = double.PositiveInfinity });
    }

    private static int CountOccurrences(string value, string needle) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(needle, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += needle.Length;
        }

        return count;
    }
}
