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
    public void Table_RejectsInvalidRelativeColumnWidthWeights() {
        var style = TableStyles.Minimal();

        var exception = Assert.Throws<ArgumentException>(() =>
            style.ColumnWidthWeights = new List<double> { 1, 0, 1 });

        Assert.Contains("Table column width weights must be positive finite values.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidFixedColumnWidthPoints_AndFitsOversizedFixedColumns() {
        var invalidStyle = TableStyles.Minimal();

        var invalidException = Assert.Throws<ArgumentException>(() =>
            invalidStyle.ColumnWidthPoints = new List<double?> { 1, -5 });

        Assert.Contains("Table fixed column widths must be positive finite values.", invalidException.Message, StringComparison.Ordinal);

        var tooWideStyle = TableStyles.Minimal();
        tooWideStyle.ColumnWidthPoints = new List<double?> { 400, 400 };

        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { "A", "B" },
                new[] { "1", "2" }
            }, style: tooWideStyle)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double firstColumnX = FindWordStartX(page, "A");
        double secondColumnX = FindWordStartX(page, "B");

        Assert.InRange(secondColumnX - firstColumnX, 210D, 240D);
    }

    [Fact]
    public void Table_RejectsInvalidMinimumAndMaximumColumnWidthPoints() {
        var invalidMinimum = TableStyles.Minimal();

        var invalidMinimumException = Assert.Throws<ArgumentException>(() =>
            invalidMinimum.ColumnMinWidthPoints = new List<double?> { 0 });

        Assert.Contains("Table minimum column widths must be positive finite values.", invalidMinimumException.Message, StringComparison.Ordinal);

        var invertedRange = TableStyles.Minimal();
        invertedRange.ColumnMinWidthPoints = new List<double?> { 90 };
        invertedRange.ColumnMaxWidthPoints = new List<double?> { 60 };

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: invertedRange)
                .ToBytes());

        var invalidHeaderRows = TableStyles.Minimal();

        var invalidHeaderRowsException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderRows.HeaderRowCount = -1);

        Assert.Contains("Table header row count cannot be negative.", invalidHeaderRowsException.Message, StringComparison.Ordinal);

        var invalidRepeatHeaderRows = TableStyles.Minimal();

        var invalidRepeatHeaderRowsException = Assert.Throws<ArgumentException>(() =>
            invalidRepeatHeaderRows.RepeatHeaderRowCount = -1);

        Assert.Contains("Table repeating header row count cannot be negative.", invalidRepeatHeaderRowsException.Message, StringComparison.Ordinal);

        var invalidFooterRows = TableStyles.Minimal();

        var invalidFooterRowsException = Assert.Throws<ArgumentException>(() =>
            invalidFooterRows.FooterRowCount = -1);

        Assert.Contains("Table footer row count cannot be negative.", invalidFooterRowsException.Message, StringComparison.Ordinal);

        var invalidMinimumRowHeight = TableStyles.Minimal();

        var invalidMinimumRowHeightException = Assert.Throws<ArgumentException>(() =>
            invalidMinimumRowHeight.MinRowHeight = -1);

        Assert.Contains("Table minimum row height must be a non-negative finite value.", invalidMinimumRowHeightException.Message, StringComparison.Ordinal);

        var invalidPreferredWidth = TableStyles.Minimal();

        var invalidPreferredWidthException = Assert.Throws<ArgumentException>(() =>
            invalidPreferredWidth.PreferredWidth = double.PositiveInfinity);

        Assert.Contains("Table preferred width must be a positive finite value.", invalidPreferredWidthException.Message, StringComparison.Ordinal);

        var invalidRowMinimumHeights = TableStyles.Minimal();

        var invalidRowMinimumHeightsException = Assert.Throws<ArgumentException>(() =>
            invalidRowMinimumHeights.RowMinHeights = new List<double?> { 18, double.NaN });

        Assert.Contains("Table row minimum heights must be non-negative finite values.", invalidRowMinimumHeightsException.Message, StringComparison.Ordinal);

        var invalidSpacingBefore = TableStyles.Minimal();

        var invalidSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            invalidSpacingBefore.SpacingBefore = -1);

        Assert.Contains("Table spacing before must be a non-negative finite value.", invalidSpacingBeforeException.Message, StringComparison.Ordinal);

        var invalidSpacingAfter = TableStyles.Minimal();

        var invalidSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            invalidSpacingAfter.SpacingAfter = double.PositiveInfinity);

        Assert.Contains("Table spacing after must be a non-negative finite value.", invalidSpacingAfterException.Message, StringComparison.Ordinal);

        var invalidCaptionFontSize = TableStyles.Minimal();
        invalidCaptionFontSize.Caption = "Caption";

        var invalidCaptionFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionFontSize.CaptionFontSize = 0);

        Assert.Contains("Table caption font size must be a positive finite value.", invalidCaptionFontSizeException.Message, StringComparison.Ordinal);

        var invalidBodyFontSize = TableStyles.Minimal();

        var invalidBodyFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidBodyFontSize.FontSize = 0);

        Assert.Contains("Table body font size must be a positive finite value.", invalidBodyFontSizeException.Message, StringComparison.Ordinal);

        var invalidLineHeight = TableStyles.Minimal();

        var invalidLineHeightException = Assert.Throws<ArgumentException>(() =>
            invalidLineHeight.LineHeight = double.NaN);

        Assert.Contains("Table line height must be a positive finite value.", invalidLineHeightException.Message, StringComparison.Ordinal);

        var invalidHeaderFontSize = TableStyles.Minimal();

        var invalidHeaderFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderFontSize.HeaderFontSize = double.NaN);

        Assert.Contains("Table header font size must be a positive finite value.", invalidHeaderFontSizeException.Message, StringComparison.Ordinal);

        var invalidFooterFontSize = TableStyles.Minimal();

        var invalidFooterFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidFooterFontSize.FooterFontSize = double.PositiveInfinity);

        Assert.Contains("Table footer font size must be a positive finite value.", invalidFooterFontSizeException.Message, StringComparison.Ordinal);

        var whitespaceCaption = TableStyles.Minimal();
        whitespaceCaption.Caption = "   ";

        var whitespaceCaptionException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: whitespaceCaption)
                .ToBytes());

        Assert.Contains("Table caption cannot be empty or whitespace.", whitespaceCaptionException.Message, StringComparison.Ordinal);

        var invalidCaptionSpacing = TableStyles.Minimal();
        invalidCaptionSpacing.Caption = "Caption";

        var invalidCaptionSpacingException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionSpacing.CaptionSpacingAfter = -1);

        Assert.Contains("Table caption spacing after must be a non-negative finite value.", invalidCaptionSpacingException.Message, StringComparison.Ordinal);

        var invalidCellBorder = TableStyles.Minimal();

        var invalidCellBorderException = Assert.Throws<ArgumentException>(() =>
            invalidCellBorder.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
                [(1, 1)] = new PdfCellBorder {
                    Width = -0.5
                }
            });

        Assert.Contains("Table cell border widths must be non-negative finite values.", invalidCellBorderException.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(-0.1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void CellBorder_RejectsInvalidWidthOnAssignment(double width) {
        var border = new PdfCellBorder();

        var exception = Assert.Throws<ArgumentException>(() =>
            border.Width = width);

        Assert.Equal("Width", exception.ParamName);
        Assert.Contains("Table cell border widths must be non-negative finite values.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidColumnSpan() {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTableCell("Invalid", 0));

        Assert.Equal("columnSpan", exception.ParamName);
        Assert.Contains("Table cell column span must be at least 1.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidRowSpan() {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTableCell("Invalid", rowSpan: 0));

        Assert.Equal("rowSpan", exception.ParamName);
        Assert.Contains("Table cell row span must be at least 1.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanBeyondAvailableRows() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { PdfTableCell.Merge("TooTall", rowSpan: 3), PdfTableCell.TextCell("A1") },
                    new[] { PdfTableCell.TextCell("A2") }
                }));

        Assert.Equal("rows", exception.ParamName);
        Assert.Contains("Table cell row span cannot extend beyond the available table rows.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanBeyondAvailableRows() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.Merge("TooTall", rowSpan: 3), PdfTableCell.TextCell("A1") },
                                        new[] { PdfTableCell.TextCell("A2") }
                                    })))))));

        Assert.Equal("rows", exception.ParamName);
        Assert.Contains("Table cell row span cannot extend beyond the available table rows.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                    new[] { PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanCrossingFooterBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { PdfTableCell.TextCell("B0"), PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.Merge("BodyFooter", rowSpan: 2), PdfTableCell.TextCell("B2") },
                    new[] { PdfTableCell.TextCell("F1") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table footer boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                                        new[] { PdfTableCell.TextCell("B1") },
                                        new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanCrossingFooterBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.TextCell("B0"), PdfTableCell.TextCell("B1") },
                                        new[] { PdfTableCell.Merge("BodyFooter", rowSpan: 2), PdfTableCell.TextCell("B2") },
                                        new[] { PdfTableCell.TextCell("F1") }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table footer boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsHeaderRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRepeatHeaderRowCountBeyondHeaderRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 2;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" },
                    new[] { "B3", "B4" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table repeating header row count cannot exceed the table header row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsCombinedHeaderAndFooterRowsBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 2;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header and footer row counts cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsFooterRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "B1", "B2" },
                                        new[] { "B3", "B4" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table footer row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsCombinedHeaderAndFooterRowsBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 2;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "H1", "H2" },
                                        new[] { "B1", "B2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table header and footer row counts cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithHeaderRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                    new[] { PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }


}
