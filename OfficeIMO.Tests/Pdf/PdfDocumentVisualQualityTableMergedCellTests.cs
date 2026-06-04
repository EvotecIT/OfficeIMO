using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Table_LinkedColumnSpanRendersAnnotationAcrossCombinedCellWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 70, 70, 60 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Span("SpannedLink", 2, "https://evotec.xyz/spanned", "Spanned cell metadata"),
                    PdfTableCell.TextCell("Tail")
                },
                new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Spanned cell metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.X2 - rect.X1 > 120, $"Expected linked spanned cell annotation to cover the combined cell width. Width: {rect.X2 - rect.X1:0.##}.");
    }

    [Fact]
    public void RowColumnTable_LinkedColumnSpanRendersAnnotationAcrossCombinedCellWidth() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 40 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Span("MergedLink", 2, "https://evotec.xyz/row-column-spanned", "Row-column spanned metadata"),
                                        PdfTableCell.TextCell("Tail")
                                    },
                                    new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column spanned metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.X2 - rect.X1 > 80, $"Expected row-column linked spanned cell annotation to cover the combined cell width. Width: {rect.X2 - rect.X1:0.##}.");
    }

    [Fact]
    public void Table_LinkedRowSpanRendersAnnotationAcrossCombinedCellHeight() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 28;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("TallLink", rowSpan: 2, linkUri: "https://evotec.xyz/row-spanned", linkContents: "Row-spanned cell metadata"),
                    PdfTableCell.TextCell("A1"),
                    PdfTableCell.TextCell("B1")
                },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-spanned cell metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.Y2 - rect.Y1 > 40, $"Expected linked row-spanned cell annotation to cover the combined cell height. Height: {rect.Y2 - rect.Y1:0.##}.");
    }

    [Fact]
    public void RowColumnTable_LinkedRowSpanRendersAnnotationAcrossCombinedCellHeight() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 28;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 40 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("TallLink", rowSpan: 2, linkUri: "https://evotec.xyz/row-column-row-spanned", linkContents: "Row-column row-spanned metadata"),
                                        PdfTableCell.TextCell("A1"),
                                        PdfTableCell.TextCell("B1")
                                    },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-row-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column row-spanned metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.Y2 - rect.Y1 > 40, $"Expected row-column linked row-spanned cell annotation to cover the combined cell height. Height: {rect.Y2 - rect.Y1:0.##}.");
    }

    [Fact]
    public void Table_RectangularMergedCellUsesCombinedBoxForFillBorderAndLink() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.23, 0.34, 0.45)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.63, 0.24, 0.14),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2, linkUri: "https://evotec.xyz/rectangular-merged", linkContents: "Rectangular merged metadata"),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.23 0.34 0.45 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.63 0.24 0.14 RG", "S"));
        var link = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.InRange(fill.W, 109, 111);
        Assert.True(fill.H > 45, $"Expected rectangular merged cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 109, 111);
        Assert.True(border.H > 45, $"Expected rectangular merged cell border to use combined row height. Height: {border.H:0.##}.");
        Assert.True(link.X2 - link.X1 > 100, $"Expected linked rectangular merged cell to cover combined width. Width: {link.X2 - link.X1:0.##}.");
        Assert.True(link.Y2 - link.Y1 >= 39, $"Expected linked rectangular merged cell to cover combined text-frame height. Height: {link.Y2 - link.Y1:0.##}.");
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/rectangular-merged)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Rectangular merged metadata)"));
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellUsesCombinedBoxForFillBorderAndLink() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.23, 0.34, 0.45)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.63, 0.24, 0.14),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2, linkUri: "https://evotec.xyz/row-column-rectangular-merged", linkContents: "Row-column rectangular merged metadata"),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.23 0.34 0.45 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.63 0.24 0.14 RG", "S"));
        var link = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.InRange(fill.W, 89, 91);
        Assert.True(fill.H > 45, $"Expected row-column rectangular merged cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 89, 91);
        Assert.True(border.H > 45, $"Expected row-column rectangular merged cell border to use combined row height. Height: {border.H:0.##}.");
        Assert.True(link.X2 - link.X1 > 80, $"Expected row-column linked rectangular merged cell to cover combined width. Width: {link.X2 - link.X1:0.##}.");
        Assert.True(link.Y2 - link.Y1 >= 39, $"Expected row-column linked rectangular merged cell to cover combined text-frame height. Height: {link.Y2 - link.Y1:0.##}.");
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-rectangular-merged)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column rectangular merged metadata)"));
    }

    [Fact]
    public void Table_RectangularMergedCellSkipsInternalVerticalGridOnContinuationRow() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var verticalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.X1 - segment.X2) < 0.01)
            .ToList();

        Assert.DoesNotContain(verticalBorders, segment => Math.Abs(segment.X1 - 80) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 172);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 80) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 148 && (segment.Y1 + segment.Y2) / 2 < 172);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 140) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 172);
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellSkipsInternalVerticalGridOnContinuationRow() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var verticalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.X1 - segment.X2) < 0.01)
            .ToList();

        Assert.DoesNotContain(verticalBorders, segment => Math.Abs(segment.X1 - 75) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 182);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 75) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 158 && (segment.Y1 + segment.Y2) / 2 < 182);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 120) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 182);
    }

    [Fact]
    public void Table_RectangularMergedCellAlignsTextInsideCombinedBox() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Center, PdfColumnAlign.Center, PdfColumnAlign.Center };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "MergedBox");
        double c2Y = FindWordStartY(page, "C2");
        double mergedY = FindWordStartY(page, "MergedBox");

        Assert.InRange(mergedX, 56, 68);
        Assert.True(Math.Abs(mergedY - c2Y) <= 3,
            $"Expected rectangular merged-cell bottom alignment to place text with the second row baseline. Merged={mergedY:0.##}, C2={c2Y:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellAlignsTextInsideCombinedBox() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Center, PdfColumnAlign.Center, PdfColumnAlign.Center };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "MergedBox");
        double c2Y = FindWordStartY(page, "C2");
        double mergedY = FindWordStartY(page, "MergedBox");

        Assert.InRange(mergedX, 46, 58);
        Assert.True(Math.Abs(mergedY - c2Y) <= 3,
            $"Expected row-column rectangular merged-cell bottom alignment to place text with the second row baseline. Merged={mergedY:0.##}, C2={c2Y:0.##}.");
    }


}
