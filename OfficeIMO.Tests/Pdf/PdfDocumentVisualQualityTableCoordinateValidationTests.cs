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
    public void Table_RejectsOutOfRangeCellFillCoordinates() {
        var style = TableStyles.Minimal();
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(2, 0)] = PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell fill coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeCellBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 2)] = new PdfCellBorder()
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell border coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeCellFillCoordinates() {
        var style = TableStyles.Minimal();
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 2)] = PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell fill coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeRowMinimumHeights() {
        var style = TableStyles.Minimal();
        style.RowMinHeights = new List<double?> {
            null,
            null,
            24
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table row minimum heights must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeRowBreakPolicies() {
        var style = TableStyles.Minimal();
        style.RowAllowBreakAcrossPages = new List<bool?> {
            null,
            null,
            false
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table row break policies must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeBodyColumnFill() {
        var style = TableStyles.Minimal();
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            null,
            PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table body column fills must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_IgnoresCellFillCoordinatesCoveredByColumnSpan() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 1)] = PdfColor.Gray
        };

        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { PdfTableCell.Span("A+B", 2), PdfTableCell.TextCell("C") },
                new[] { PdfTableCell.TextCell("1"), PdfTableCell.TextCell("2"), PdfTableCell.TextCell("3") }
            }, style: style)
            .ToBytes();

        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes));
    }

    [Fact]
    public void Table_IgnoresCellAlignmentCoordinatesCoveredByRowSpan() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
            [(1, 0)] = PdfColumnAlign.Right
        };

        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { PdfTableCell.Merge("A", rowSpan: 2), PdfTableCell.TextCell("B") },
                new[] { PdfTableCell.TextCell("C") }
            }, style: style)
            .ToBytes();

        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes));
    }

    [Fact]
    public void Table_RejectsOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeColumnWidthWeight() {
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> {
            1,
            1,
            1
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table column width weights must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeFixedColumnWidth() {
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> {
            null,
            null,
            42
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table fixed column widths must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidLinkUri() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfTableCell.TextCell("Invalid", "bad\u0001uri"));

        Assert.Equal("linkUri", exception.ParamName);
        Assert.Contains("Parameter 'linkUri' must be an absolute URI or a relative URI action target.", exception.Message, StringComparison.Ordinal);
    }


}
