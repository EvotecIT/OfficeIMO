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
    public void Table_ColumnSpanUsesCombinedColumnWidthAndSnapshotsCells() {
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
        var spanned = PdfTableCell.Span("SpannedTitle", 2);
        var rows = new[] {
            new[] { spanned, PdfTableCell.TextCell("Tail") },
            new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
        };

        PdfDocument doc = PdfDocument.Create(options)
            .Table(rows, style: style);

        rows[0][0] = PdfTableCell.TextCell("Mutated");
        byte[] bytes = doc.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double spannedX = FindWordStartX(page, "SpannedTitle");
        double tailX = FindWordStartX(page, "Tail");
        double bX = FindWordStartX(page, "B");
        double cX = FindWordStartX(page, "C");

        Assert.InRange(spannedX, 33, 38);
        Assert.InRange(tailX, 173, 178);
        Assert.InRange(bX, 103, 108);
        Assert.InRange(cX, 173, 178);
        Assert.DoesNotContain("Mutated", page.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_ColumnSpanUsesCombinedColumnWidth() {
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
                                    new[] { PdfTableCell.Span("Merged", 2), PdfTableCell.TextCell("Tail") },
                                    new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "Merged");
        double tailX = FindWordStartX(page, "Tail");
        double bX = FindWordStartX(page, "B");
        double cX = FindWordStartX(page, "C");

        Assert.InRange(mergedX, 33, 38);
        Assert.InRange(tailX, 133, 138);
        Assert.InRange(bX, 83, 88);
        Assert.InRange(cX, 133, 138);
    }

    [Fact]
    public void Table_RowSpanOccupiesFollowingRowGridColumn() {
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
        style.CellPaddingY = 4;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Middle };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { PdfTableCell.Merge("GroupOne", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("Tail0"), PdfTableCell.TextCell("Tail1"), PdfTableCell.TextCell("Tail2") }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double groupX = FindWordStartX(page, "GroupOne");
        double a1X = FindWordStartX(page, "A1");
        double b1X = FindWordStartX(page, "B1");
        double a2X = FindWordStartX(page, "A2");
        double b2X = FindWordStartX(page, "B2");
        double groupY = FindWordStartY(page, "GroupOne");
        double a1Y = FindWordStartY(page, "A1");
        double a2Y = FindWordStartY(page, "A2");

        Assert.InRange(groupX, 33, 38);
        Assert.InRange(a1X, 93, 99);
        Assert.InRange(b1X, 163, 169);
        Assert.InRange(a2X, 93, 99);
        Assert.InRange(b2X, 163, 169);
        Assert.True(groupY < a1Y - 2 && groupY > a2Y + 2,
            $"Expected vertically centered row-spanned text between row baselines. Group={groupY:0.##}, A1={a1Y:0.##}, A2={a2Y:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RowSpanOccupiesFollowingRowGridColumn() {
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
        style.CellPaddingY = 4;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Middle };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Merge", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "Merge");
        double a1X = FindWordStartX(page, "A1");
        double b1X = FindWordStartX(page, "B1");
        double a2X = FindWordStartX(page, "A2");
        double b2X = FindWordStartX(page, "B2");
        double mergedY = FindWordStartY(page, "Merge");
        double a1Y = FindWordStartY(page, "A1");
        double a2Y = FindWordStartY(page, "A2");

        Assert.InRange(mergedX, 33, 38);
        Assert.InRange(a1X, 78, 84);
        Assert.InRange(b1X, 123, 129);
        Assert.InRange(a2X, 78, 84);
        Assert.InRange(b2X, 123, 129);
        Assert.True(mergedY < a1Y - 2 && mergedY > a2Y + 2,
            $"Expected row-column row-spanned text between row baselines. Merged={mergedY:0.##}, A1={a1Y:0.##}, A2={a2Y:0.##}.");
    }

    [Fact]
    public void Table_RowSpanCellFillAndBorderUseCombinedRowHeight() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));

        Assert.InRange(fill.W, 59, 61);
        Assert.True(fill.H > 45, $"Expected row-spanned cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 59, 61);
        Assert.True(border.H > 45, $"Expected row-spanned cell border to use combined row height. Height: {border.H:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RowSpanCellFillAndBorderUseCombinedRowHeight() {
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
            [(0, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));

        Assert.InRange(fill.W, 44, 46);
        Assert.True(fill.H > 45, $"Expected row-column row-spanned cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 44, 46);
        Assert.True(border.H > 45, $"Expected row-column row-spanned cell border to use combined row height. Height: {border.H:0.##}.");
    }

    [Fact]
    public void Table_RowSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void RowColumnTable_RowSpanIgnoresContinuationCellFillAndBorderCoordinates() {
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
            [(1, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void Table_ColumnSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 1)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void RowColumnTable_ColumnSpanIgnoresContinuationCellFillAndBorderCoordinates() {
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
            [(0, 1)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }


}
