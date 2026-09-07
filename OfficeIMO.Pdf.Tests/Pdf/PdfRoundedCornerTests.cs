using System;
using System.Collections.Generic;
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
    public void CornerRadius_ClonesAndRejectsNegative() {
        PdfTableStyle table = new() { CornerRadius = 6 };
        Assert.Equal(6, table.Clone().CornerRadius);
        Assert.Throws<ArgumentException>(() => new PdfTableStyle { CornerRadius = -1 });

        PdfPanelStyle panel = new() { CornerRadius = 8 };
        Assert.Equal(8, panel.Clone().CornerRadius);
        Assert.Throws<ArgumentException>(() => new PdfPanelStyle { CornerRadius = -1 });
    }
}
