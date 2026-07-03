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
    public void Table_RowSpanSkipsContinuationRowStripeFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

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
        var stripe = Assert.Single(ExtractPaintedRectangles(content, "0.21 0.31 0.41 rg", "f"));

        Assert.InRange(stripe.X, 89, 91);
        Assert.InRange(stripe.W, 139, 141);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsContinuationRowStripeFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

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
        var stripe = Assert.Single(ExtractPaintedRectangles(content, "0.21 0.31 0.41 rg", "f"));

        Assert.InRange(stripe.X, 74, 76);
        Assert.InRange(stripe.W, 89, 91);
    }

    [Fact]
    public void Table_RowSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
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
        style.BodyColumnFills = new List<PdfColor?> {
            new PdfColor(0.11, 0.22, 0.33)
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
        var fills = ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f");

        Assert.Equal(2, fills.Count);
        Assert.All(fills, fill => {
            Assert.InRange(fill.X, 29, 31);
            Assert.InRange(fill.W, 59, 61);
        });
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
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
        style.BodyColumnFills = new List<PdfColor?> {
            new PdfColor(0.11, 0.22, 0.33)
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
        var fills = ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f");

        Assert.Equal(2, fills.Count);
        Assert.All(fills, fill => {
            Assert.InRange(fill.X, 29, 31);
            Assert.InRange(fill.W, 44, 46);
        });
    }

    [Fact]
    public void Table_ColumnSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
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
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
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
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f"));

        Assert.InRange(fill.X, 89, 91);
        Assert.InRange(fill.W, 69, 71);
    }

    [Fact]
    public void RowColumnTable_ColumnSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
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
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
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
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f"));

        Assert.InRange(fill.X, 74, 76);
        Assert.InRange(fill.W, 44, 46);
    }

    [Fact]
    public void Table_RowSpanSkipsInternalRowSeparatorAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

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
        var separators = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG");
        var partial = Assert.Single(separators, segment => segment.X1 > 88 && segment.X1 < 92 && segment.X2 > 228 && segment.X2 < 232);

        Assert.Contains(separators, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 228 && segment.X2 < 232);
        Assert.DoesNotContain(separators, segment => Math.Abs(segment.Y1 - partial.Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsInternalRowSeparatorAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

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
        var separators = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG");
        var partial = Assert.Single(separators, segment => segment.X1 > 73 && segment.X1 < 77 && segment.X2 > 163 && segment.X2 < 167);

        Assert.Contains(separators, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 163 && segment.X2 < 167);
        Assert.DoesNotContain(separators, segment => Math.Abs(segment.Y1 - partial.Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
    }

    [Fact]
    public void Table_RowSpanSkipsInternalBorderLineAcrossMergedCell() {
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
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var horizontalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.Y1 - segment.Y2) < 0.01)
            .ToList();
        var partials = horizontalBorders
            .Where(segment => segment.X1 > 88 && segment.X1 < 92 && segment.X2 > 228 && segment.X2 < 232)
            .ToList();

        Assert.NotEmpty(partials);
        Assert.Contains(horizontalBorders, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 228 && segment.X2 < 232);
        Assert.DoesNotContain(horizontalBorders, segment => Math.Abs(segment.Y1 - partials[0].Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsInternalBorderLineAcrossMergedCell() {
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

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 180,
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
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var horizontalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.Y1 - segment.Y2) < 0.01)
            .ToList();
        var partials = horizontalBorders
            .Where(segment => segment.X1 > 73 && segment.X1 < 77 && segment.X2 > 163 && segment.X2 < 167)
            .ToList();

        Assert.NotEmpty(partials);
        Assert.Contains(horizontalBorders, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 163 && segment.X2 < 167);
        Assert.DoesNotContain(horizontalBorders, segment => Math.Abs(segment.Y1 - partials[0].Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
        Assert.DoesNotContain(" re S", content);
    }



}
