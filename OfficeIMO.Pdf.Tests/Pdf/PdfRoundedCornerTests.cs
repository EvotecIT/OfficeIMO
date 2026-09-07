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
                new PdfTextRun("Callout title", bold: true, color: new PdfColor(0.1, 0.1, 0.1), fontSize: 11),
                new PdfTextRun("\n", color: new PdfColor(0.1, 0.1, 0.1), fontSize: 5),
                new PdfTextRun("Body with ", color: new PdfColor(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun("bold", bold: true, color: new PdfColor(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun(" and a ", color: new PdfColor(0.2, 0.2, 0.2), fontSize: 10),
                new PdfTextRun("link", color: new PdfColor(0.1, 0.4, 0.7), fontSize: 10, linkUri: "https://example.com"),
                new PdfTextRun(".", color: new PdfColor(0.2, 0.2, 0.2), fontSize: 10),
            }),
        },
    };

    private static PdfTableStyle CalloutStyle(double radius) => new() {
        HeaderRowCount = 0,
        BorderColor = new PdfColor(0.85, 0.85, 0.85),
        BorderWidth = 1,
        RowSeparatorWidth = 0,
        CornerRadius = radius,
        CellFills = new Dictionary<(int, int), PdfColor> { [(0, 0)] = new PdfColor(0.95, 0.97, 0.99) },
        CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder { LeftBorder = new PdfCellBorderSide { Color = new PdfColor(0.13, 0.47, 0.71), Width = 4 } },
        },
    };

    [Fact]
    public void TableCornerRadius_RoundsOuterBox_ClipsCellBorder_AndKeepsInlineText() {
        byte[] rounded = PdfDocument.Create(d => d.Content(c => c.Table(Callout(), PdfAlign.Left, CalloutStyle(6)))).ToBytes();
        byte[] square = PdfDocument.Create(d => d.Content(c => c.Table(Callout(), PdfAlign.Left, CalloutStyle(0)))).ToBytes();

        string roundedRaw = PdfEncoding.Latin1GetString(rounded);
        string squareRaw = PdfEncoding.Latin1GetString(square);

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
                Background = new PdfColor(0.95, 0.97, 0.99),
                BorderColor = new PdfColor(0.2, 0.3, 0.4),
                BorderWidth = 1,
                CornerRadius = 8,
                PaddingX = 12,
                PaddingY = 10,
            })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        Assert.Contains(" c", raw);   // rounded corners emit bezier operators
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Panel title", text, StringComparison.Ordinal);
        Assert.Contains("Panel body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTableCornerRadius_RoundsTheFixedPositionOuterBox() {
        var style = CalloutStyle(8);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(Callout(), 24, 24, 160, 64, style))
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.Contains("Callout title", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void DeferredTableCornerRadius_UsesLogicalEdgesRegardlessOfBatchSize() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CornerRadius = 7;

        byte[] oneRowBatches = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TableDeferred(CreateRows, batchSize: 1, style: style)
            .ToBytes();
        byte[] threeRowBatches = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TableDeferred(CreateRows, batchSize: 3, style: style)
            .ToBytes();

        Assert.Equal(
            CountOccurrences(PdfEncoding.Latin1GetString(oneRowBatches), " c"),
            CountOccurrences(PdfEncoding.Latin1GetString(threeRowBatches), " c"));

        static IEnumerable<string[]> CreateRows() {
            for (int row = 0; row < 7; row++) yield return new[] { "Row " + row, "Value " + row };
        }
    }

    [Fact]
    public void RoundedCellBorder_PreservesDashDoubleLineAndDiagonals() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.CornerRadius = 8;
        style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(68, 85, 102),
                Width = 1,
                DashStyle = OfficeStrokeDashStyle.Dash,
                LineStyle = PdfCellBorderLineStyle.TwoLine,
                DiagonalUp = true,
                DiagonalDown = true
            }
        };

        string raw = PdfEncoding.Latin1GetString(PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Table(new[] { new[] { "Border contract" } }, style: style)
            .ToBytes());

        Assert.Contains("[3 1.5] 0 d", raw, StringComparison.Ordinal);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.True(CountOccurrences(raw, " S") >= 12, "Expected two strokes for every side and both diagonals.");
    }

    [Fact]
    public void CornerRadius_DoesNotRoundInteriorCellCorners() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.FooterFill = null;
        style.RowStripeFill = null;
        style.CornerRadius = 8;
        style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder { Color = PdfColor.FromRgb(68, 85, 102), Width = 1 }
        };

        string raw = PdfEncoding.Latin1GetString(PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Table(new[] { new[] { "Left", "Middle", "Right" } }, style: style)
            .ToBytes());
        style.CellBorders = null;
        string withoutCellBorder = PdfEncoding.Latin1GetString(PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
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
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.CornerRadius = 10;
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.ColumnWidthPoints = new List<double?> { 60, 60 };
        style.FixedRowHeights = new List<double?> { 30, 30 };
        style.CellFills = new Dictionary<(int, int), PdfColor> { [(0, 0)] = PdfColor.FromRgb(255, 0, 0) };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 160,
                PageHeight = 120,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
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

        byte[] expectedFlow = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Table(rows, style: CreateStyle(10))
            .ToBytes();
        byte[] oversizedFlow = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Table(rows, style: CreateStyle(100))
            .ToBytes();
        byte[] expectedCanvas = PdfDocument.Create(new PdfOptions { PageWidth = 180, PageHeight = 120, MarginLeft = 10, MarginRight = 10, MarginTop = 10, MarginBottom = 10, CompressContentStreams = false })
            .Canvas(canvas => canvas.Table(rows, 20, 20, 120, 60, CreateStyle(10)))
            .ToBytes();
        byte[] oversizedCanvas = PdfDocument.Create(new PdfOptions { PageWidth = 180, PageHeight = 120, MarginLeft = 10, MarginRight = 10, MarginTop = 10, MarginBottom = 10, CompressContentStreams = false })
            .Canvas(canvas => canvas.Table(rows, 20, 20, 120, 60, CreateStyle(100)))
            .ToBytes();

        Assert.Equal(expectedFlow, oversizedFlow);
        Assert.Equal(expectedCanvas, oversizedCanvas);

        static PdfTableStyle CreateStyle(double radius) {
            PdfTableStyle style = TableStyles.Minimal();
            style.HeaderRowCount = 0;
            style.BorderColor = PdfColor.FromRgb(68, 85, 102);
            style.BorderWidth = 1;
            style.CornerRadius = radius;
            style.CellPaddingX = 0;
            style.CellPaddingY = 0;
            style.ColumnWidthPoints = new List<double?> { 60, 60 };
            style.FixedRowHeights = new List<double?> { 20, 40 };
            style.CellFills = new Dictionary<(int, int), PdfColor> { [(0, 0)] = PdfColor.FromRgb(220, 235, 250) };
            style.CellBorders = new Dictionary<(int, int), PdfCellBorder> {
                [(0, 0)] = new PdfCellBorder { Color = PdfColor.FromRgb(24, 119, 181), Width = 3 }
            };
            return style;
        }
    }

    [Fact]
    public void CornerRadius_ClonesAndRejectsNegative() {
        PdfTableStyle table = new() { CornerRadius = 6 };
        Assert.Equal(6, table.Clone().CornerRadius);
        Assert.Throws<ArgumentException>(() => new PdfTableStyle { CornerRadius = -1 });
        Assert.Throws<ArgumentException>(() => new PdfTableStyle { CornerRadius = double.NaN });
        Assert.Throws<ArgumentException>(() => new PdfTableStyle { CornerRadius = double.PositiveInfinity });

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
