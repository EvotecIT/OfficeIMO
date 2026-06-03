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
    public void Table_RejectsInvalidStylePrimitiveValues() {
        var invalidBorder = TableStyles.Minimal();

        var borderException = Assert.Throws<ArgumentException>(() =>
            invalidBorder.BorderWidth = double.NaN);

        Assert.Contains("Table border width must be a non-negative finite value.", borderException.Message, StringComparison.Ordinal);

        var invalidRowSeparator = TableStyles.Minimal();

        var rowSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidRowSeparator.RowSeparatorWidth = double.PositiveInfinity);

        Assert.Contains("Table row separator width must be a non-negative finite value.", rowSeparatorException.Message, StringComparison.Ordinal);

        var invalidHeaderSeparator = TableStyles.Minimal();

        var headerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderSeparator.HeaderSeparatorWidth = -0.1);

        Assert.Contains("Table header separator width must be a non-negative finite value.", headerSeparatorException.Message, StringComparison.Ordinal);

        var invalidFooterSeparator = TableStyles.Minimal();

        var footerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidFooterSeparator.FooterSeparatorWidth = -0.1);

        Assert.Contains("Table footer separator width must be a non-negative finite value.", footerSeparatorException.Message, StringComparison.Ordinal);

        var invalidMaxWidth = TableStyles.Minimal();

        var maxWidthException = Assert.Throws<ArgumentException>(() =>
            invalidMaxWidth.MaxWidth = 0);

        Assert.Contains("Table max width must be a positive finite value.", maxWidthException.Message, StringComparison.Ordinal);

        var invalidLeftIndent = TableStyles.Minimal();

        var leftIndentException = Assert.Throws<ArgumentException>(() =>
            invalidLeftIndent.LeftIndent = -1);

        Assert.Contains("Table left indent must be a non-negative finite value.", leftIndentException.Message, StringComparison.Ordinal);

        var invalidHorizontalPadding = TableStyles.Minimal();

        var horizontalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidHorizontalPadding.CellPaddingX = -1);

        Assert.Contains("Table horizontal cell padding must be a non-negative finite value.", horizontalPaddingException.Message, StringComparison.Ordinal);

        var invalidVerticalPadding = TableStyles.Minimal();

        var verticalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalPadding.CellPaddingY = double.PositiveInfinity);

        Assert.Contains("Table vertical cell padding must be a non-negative finite value.", verticalPaddingException.Message, StringComparison.Ordinal);

        var invalidLeftPadding = TableStyles.Minimal();

        var leftPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidLeftPadding.CellPaddingLeft = -1);

        Assert.Contains("Table left cell padding must be a non-negative finite value.", leftPaddingException.Message, StringComparison.Ordinal);

        var invalidRightPadding = TableStyles.Minimal();

        var rightPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidRightPadding.CellPaddingRight = double.NaN);

        Assert.Contains("Table right cell padding must be a non-negative finite value.", rightPaddingException.Message, StringComparison.Ordinal);

        var invalidTopPadding = TableStyles.Minimal();

        var topPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidTopPadding.CellPaddingTop = double.PositiveInfinity);

        Assert.Contains("Table top cell padding must be a non-negative finite value.", topPaddingException.Message, StringComparison.Ordinal);

        var invalidBottomPadding = TableStyles.Minimal();

        var bottomPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidBottomPadding.CellPaddingBottom = -0.1);

        Assert.Contains("Table bottom cell padding must be a non-negative finite value.", bottomPaddingException.Message, StringComparison.Ordinal);

        var invalidCellSpacing = TableStyles.Minimal();

        var cellSpacingException = Assert.Throws<ArgumentException>(() =>
            invalidCellSpacing.CellSpacing = -1);

        Assert.Contains("Table cell spacing must be a non-negative finite value.", cellSpacingException.Message, StringComparison.Ordinal);

        var excessiveHorizontalPadding = TableStyles.Minimal();
        excessiveHorizontalPadding.ColumnWidthPoints = new List<double?> { 12 };
        excessiveHorizontalPadding.CellPaddingX = 6;

        var textWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A" },
                    new[] { "1" }
                }, style: excessiveHorizontalPadding)
                .ToBytes());

        Assert.Contains("Table horizontal cell padding must leave a positive text width.", textWidthException.Message, StringComparison.Ordinal);

        var excessiveLeftIndent = TableStyles.Minimal();
        excessiveLeftIndent.LeftIndent = 400;

        var leftIndentWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 180,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30
                })
                .Table(new[] {
                    new[] { "Only" }
                }, style: excessiveLeftIndent)
                .ToBytes());

        Assert.Contains("Table left indent must leave a positive table width.", leftIndentWidthException.Message, StringComparison.Ordinal);

        var invalidBaselineOffset = TableStyles.Minimal();

        var baselineException = Assert.Throws<ArgumentException>(() =>
            invalidBaselineOffset.RowBaselineOffset = double.NaN);

        Assert.Contains("Table row baseline offset must be a finite value.", baselineException.Message, StringComparison.Ordinal);

        var invalidCellFill = TableStyles.Minimal();

        var fillException = Assert.Throws<ArgumentException>(() =>
            invalidCellFill.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
                [(-1, 0)] = PdfColor.Gray
            });

        Assert.Contains("Table cell fill coordinates cannot be negative.", fillException.Message, StringComparison.Ordinal);

        var invalidCellDataBar = TableStyles.Minimal();

        var dataBarCoordinateException = Assert.Throws<ArgumentException>(() =>
            invalidCellDataBar.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
                [(-1, 0)] = new PdfCellDataBar { Ratio = 0.5 }
            });

        Assert.Contains("Table cell data bar coordinates cannot be negative.", dataBarCoordinateException.Message, StringComparison.Ordinal);

        var dataBarRatioException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfCellDataBar { Ratio = double.NaN });

        Assert.Contains("PDF table data bar ratio must be a finite value between 0 and 1.", dataBarRatioException.Message, StringComparison.Ordinal);

        var invalidCellIcon = TableStyles.Minimal();

        var iconCoordinateException = Assert.Throws<ArgumentException>(() =>
            invalidCellIcon.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
                [(-1, 0)] = new PdfCellIcon()
            });

        Assert.Contains("Table cell icon coordinates cannot be negative.", iconCoordinateException.Message, StringComparison.Ordinal);

        var iconSizeException = Assert.Throws<ArgumentException>(() =>
            new PdfCellIcon { Size = double.PositiveInfinity });

        Assert.Contains("PDF table cell icon size must be a positive finite value.", iconSizeException.Message, StringComparison.Ordinal);

        var iconOffsetException = Assert.Throws<ArgumentException>(() =>
            new PdfCellIcon { OffsetY = double.NaN });

        Assert.Contains("PDF table cell icon offsets must be finite values.", iconOffsetException.Message, StringComparison.Ordinal);

        var invalidCellPadding = TableStyles.Minimal();

        var paddingException = Assert.Throws<ArgumentException>(() =>
            invalidCellPadding.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
                [(-1, 0)] = new PdfCellPadding { Left = 4 }
            });

        Assert.Contains("Table cell padding coordinates cannot be negative.", paddingException.Message, StringComparison.Ordinal);

        var invalidCellPaddingValueException = Assert.Throws<ArgumentException>(() =>
            new PdfCellPadding { Left = double.NaN });

        Assert.Contains("Table cell padding values must be non-negative finite values.", invalidCellPaddingValueException.Message, StringComparison.Ordinal);

        var invalidCellAlignment = TableStyles.Minimal();

        var cellAlignmentException = Assert.Throws<ArgumentException>(() =>
            invalidCellAlignment.CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(-1, 0)] = PdfColumnAlign.Center
            });

        Assert.Contains("Table cell alignment coordinates cannot be negative.", cellAlignmentException.Message, StringComparison.Ordinal);

        var invalidCellVerticalAlignment = TableStyles.Minimal();

        var cellVerticalAlignmentException = Assert.Throws<ArgumentException>(() =>
            invalidCellVerticalAlignment.CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(-1, 0)] = PdfCellVerticalAlign.Middle
            });

        Assert.Contains("Table cell vertical alignment coordinates cannot be negative.", cellVerticalAlignmentException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidAlignmentEnumValues() {
        var invalidTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, align: (PdfAlign)99)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", invalidTableAlignException.Message, StringComparison.Ordinal);

        var unsupportedTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, align: PdfAlign.Justify)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", unsupportedTableAlignException.Message, StringComparison.Ordinal);

        var composeTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Table(new[] {
                                    new[] { "A", "B" },
                                    new[] { "1", "2" }
                                }, align: (PdfAlign)99)))))
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", composeTableAlignException.Message, StringComparison.Ordinal);

        var tableWithLinksAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .TableWithLinks(
                    new[] {
                        new[] { "A", "B" },
                        new[] { "1", "2" }
                    },
                    new Dictionary<(int Row, int Col), string> {
                        [(1, 0)] = "https://evotec.xyz"
                    },
                    align: PdfAlign.Justify));

        Assert.Contains("Table alignment must be Left, Center, or Right.", tableWithLinksAlignException.Message, StringComparison.Ordinal);

        var invalidCaptionAlign = TableStyles.Minimal();
        invalidCaptionAlign.Caption = "Caption";

        var invalidCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionAlign.CaptionAlign = (PdfAlign)99);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", invalidCaptionAlignException.Message, StringComparison.Ordinal);

        var unsupportedCaptionAlign = TableStyles.Minimal();
        unsupportedCaptionAlign.Caption = "Caption";

        var unsupportedCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            unsupportedCaptionAlign.CaptionAlign = PdfAlign.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", unsupportedCaptionAlignException.Message, StringComparison.Ordinal);

        var invalidColumnAlign = TableStyles.Minimal();

        var invalidColumnAlignException = Assert.Throws<ArgumentException>(() =>
            invalidColumnAlign.Alignments = new List<PdfColumnAlign> { (PdfColumnAlign)99 });

        Assert.Contains("Table column alignments must be Left, Center, or Right.", invalidColumnAlignException.Message, StringComparison.Ordinal);

        var invalidVerticalAlign = TableStyles.Minimal();

        var invalidVerticalAlignException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalAlign.VerticalAlignments = new List<PdfCellVerticalAlign> { (PdfCellVerticalAlign)99 });

        Assert.Contains("Table vertical alignments must be defined PDF cell vertical alignment values.", invalidVerticalAlignException.Message, StringComparison.Ordinal);

        var invalidCellAlign = TableStyles.Minimal();

        var invalidCellAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCellAlign.CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(0, 0)] = (PdfColumnAlign)99
            });

        Assert.Contains("Table column alignments must be Left, Center, or Right.", invalidCellAlignException.Message, StringComparison.Ordinal);

        var invalidCellVerticalAlign = TableStyles.Minimal();

        var invalidCellVerticalAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCellVerticalAlign.CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(0, 0)] = (PdfCellVerticalAlign)99
            });

        Assert.Contains("Table vertical alignments must be defined PDF cell vertical alignment values.", invalidCellVerticalAlignException.Message, StringComparison.Ordinal);
    }



}
