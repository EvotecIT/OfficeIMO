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
        var invalidBorder = PdfTableStyles.Minimal();

        var borderException = Assert.Throws<ArgumentException>(() =>
            invalidBorder.BorderWidth = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table border width must be a non-negative finite value.", borderException.Message, StringComparison.Ordinal);

        var invalidRowSeparator = PdfTableStyles.Minimal();

        var rowSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidRowSeparator.RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table row separator width must be a non-negative finite value.", rowSeparatorException.Message, StringComparison.Ordinal);

        var invalidHeaderSeparator = PdfTableStyles.Minimal();

        var headerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderSeparator.HeaderSeparatorWidth = global::OfficeIMO.OfficeLength.Points(-0.1));

        Assert.Contains("Table header separator width must be a non-negative finite value.", headerSeparatorException.Message, StringComparison.Ordinal);

        var invalidFooterSeparator = PdfTableStyles.Minimal();

        var footerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidFooterSeparator.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(-0.1));

        Assert.Contains("Table footer separator width must be a non-negative finite value.", footerSeparatorException.Message, StringComparison.Ordinal);

        var invalidMaxWidth = PdfTableStyles.Minimal();

        var maxWidthException = Assert.Throws<ArgumentException>(() =>
            invalidMaxWidth.MaxWidth = global::OfficeIMO.OfficeLength.Points(0));

        Assert.Contains("Table max width must be a positive finite value.", maxWidthException.Message, StringComparison.Ordinal);

        var invalidHorizontalPadding = PdfTableStyles.Minimal();

        var horizontalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidHorizontalPadding.CellPaddingX = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table horizontal cell padding must be a non-negative finite value.", horizontalPaddingException.Message, StringComparison.Ordinal);

        var invalidVerticalPadding = PdfTableStyles.Minimal();

        var verticalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalPadding.CellPaddingY = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table vertical cell padding must be a non-negative finite value.", verticalPaddingException.Message, StringComparison.Ordinal);

        var invalidLeftPadding = PdfTableStyles.Minimal();

        var leftPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidLeftPadding.CellPaddingLeft = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table left cell padding must be a non-negative finite value.", leftPaddingException.Message, StringComparison.Ordinal);

        var invalidRightPadding = PdfTableStyles.Minimal();

        var rightPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidRightPadding.CellPaddingRight = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table right cell padding must be a non-negative finite value.", rightPaddingException.Message, StringComparison.Ordinal);

        var invalidTopPadding = PdfTableStyles.Minimal();

        var topPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidTopPadding.CellPaddingTop = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table top cell padding must be a non-negative finite value.", topPaddingException.Message, StringComparison.Ordinal);

        var invalidBottomPadding = PdfTableStyles.Minimal();

        var bottomPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidBottomPadding.CellPaddingBottom = global::OfficeIMO.OfficeLength.Points(-0.1));

        Assert.Contains("Table bottom cell padding must be a non-negative finite value.", bottomPaddingException.Message, StringComparison.Ordinal);

        var invalidCellSpacing = PdfTableStyles.Minimal();

        var cellSpacingException = Assert.Throws<ArgumentException>(() =>
            invalidCellSpacing.CellSpacing = global::OfficeIMO.OfficeLength.Points(-1));

        Assert.Contains("Table cell spacing must be a non-negative finite value.", cellSpacingException.Message, StringComparison.Ordinal);

        var excessiveHorizontalPadding = PdfTableStyles.Minimal();
        excessiveHorizontalPadding.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(12) };
        excessiveHorizontalPadding.CellPaddingX = global::OfficeIMO.OfficeLength.Points(6);

        var textWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A" },
                    new[] { "1" }
                }, style: excessiveHorizontalPadding)
                .ToBytes());

        Assert.Contains("Table horizontal cell padding must leave a positive text width.", textWidthException.Message, StringComparison.Ordinal);

        var excessiveLeftIndent = PdfTableStyles.Minimal();
        excessiveLeftIndent.LeftIndent = global::OfficeIMO.OfficeLength.Points(400);

        var leftIndentWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfCreateOptions {
                    PageWidth = global::OfficeIMO.OfficeLength.Points(180),
                    PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                    MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                    MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                    MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                    MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
                })
                .Table(new[] {
                    new[] { "Only" }
                }, style: excessiveLeftIndent)
                .ToBytes());

        Assert.Contains("Table left indent must leave a positive table width.", leftIndentWidthException.Message, StringComparison.Ordinal);

        var invalidCellFill = PdfTableStyles.Minimal();

        var fillException = Assert.Throws<ArgumentException>(() =>
            invalidCellFill.CellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor> {
                [(-1, 0)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D)
            });

        Assert.Contains("Table cell fill coordinates cannot be negative.", fillException.Message, StringComparison.Ordinal);

        var invalidCellDataBar = PdfTableStyles.Minimal();

        var dataBarCoordinateException = Assert.Throws<ArgumentException>(() =>
            invalidCellDataBar.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
                [(-1, 0)] = new PdfCellDataBar { Ratio = 0.5 }
            });

        Assert.Contains("Table cell data bar coordinates cannot be negative.", dataBarCoordinateException.Message, StringComparison.Ordinal);

        var dataBarRatioException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfCellDataBar { Ratio = double.NaN });

        Assert.Contains("PDF table data bar ratio must be a finite value between 0 and 1.", dataBarRatioException.Message, StringComparison.Ordinal);

        var invalidCellIcon = PdfTableStyles.Minimal();

        var iconCoordinateException = Assert.Throws<ArgumentException>(() =>
            invalidCellIcon.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
                [(-1, 0)] = new PdfCellIcon()
            });

        Assert.Contains("Table cell icon coordinates cannot be negative.", iconCoordinateException.Message, StringComparison.Ordinal);

        var iconSizeException = Assert.Throws<ArgumentException>(() =>
            new PdfCellIcon { Size = global::OfficeIMO.OfficeLength.Points(-1) });

        Assert.Contains("PDF table cell icon size must be a positive finite value.", iconSizeException.Message, StringComparison.Ordinal);

        var invalidCellPadding = PdfTableStyles.Minimal();

        var paddingException = Assert.Throws<ArgumentException>(() =>
            invalidCellPadding.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
                [(-1, 0)] = new PdfCellPadding { Left = global::OfficeIMO.OfficeLength.Points(4) }
            });

        Assert.Contains("Table cell padding coordinates cannot be negative.", paddingException.Message, StringComparison.Ordinal);

        var invalidCellPaddingValueException = Assert.Throws<ArgumentException>(() =>
            new PdfCellPadding { Left = global::OfficeIMO.OfficeLength.Points(-1) });

        Assert.Contains("Table cell padding values must be non-negative finite values.", invalidCellPaddingValueException.Message, StringComparison.Ordinal);

        var invalidCellAlignment = PdfTableStyles.Minimal();

        var cellAlignmentException = Assert.Throws<ArgumentException>(() =>
            invalidCellAlignment.CellAlignments = new Dictionary<(int Row, int Column), OfficeAlignment> {
                [(-1, 0)] = OfficeAlignment.Center
            });

        Assert.Contains("Table cell alignment coordinates cannot be negative.", cellAlignmentException.Message, StringComparison.Ordinal);

        var invalidCellVerticalAlignment = PdfTableStyles.Minimal();

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
                }, align: (OfficeAlignment)99)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, Right, Start, or End.", invalidTableAlignException.Message, StringComparison.Ordinal);

        var unsupportedTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, align: OfficeAlignment.Justify)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, Right, Start, or End.", unsupportedTableAlignException.Message, StringComparison.Ordinal);

        var composeTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Table(new[] {
                                    new[] { "A", "B" },
                                    new[] { "1", "2" }
                                }, align: (OfficeAlignment)99)))))
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, Right, Start, or End.", composeTableAlignException.Message, StringComparison.Ordinal);

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
                    align: OfficeAlignment.Justify));

        Assert.Contains("Table alignment must be Left, Center, Right, Start, or End.", tableWithLinksAlignException.Message, StringComparison.Ordinal);

        var invalidCaptionAlign = PdfTableStyles.Minimal();
        invalidCaptionAlign.Caption = "Caption";

        var invalidCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionAlign.CaptionAlign = (OfficeAlignment)99);

        Assert.Contains("Table caption alignment must be Left, Center, Right, Start, or End.", invalidCaptionAlignException.Message, StringComparison.Ordinal);

        var unsupportedCaptionAlign = PdfTableStyles.Minimal();
        unsupportedCaptionAlign.Caption = "Caption";

        var unsupportedCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            unsupportedCaptionAlign.CaptionAlign = OfficeAlignment.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, Right, Start, or End.", unsupportedCaptionAlignException.Message, StringComparison.Ordinal);

        var invalidColumnAlign = PdfTableStyles.Minimal();

        var invalidColumnAlignException = Assert.Throws<ArgumentException>(() =>
            invalidColumnAlign.Alignments = new List<OfficeAlignment> { (OfficeAlignment)99 });

        Assert.Contains("Table column alignments must be Left, Center, Right, Start, or End.", invalidColumnAlignException.Message, StringComparison.Ordinal);

        var invalidVerticalAlign = PdfTableStyles.Minimal();

        var invalidVerticalAlignException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalAlign.VerticalAlignments = new List<PdfCellVerticalAlign> { (PdfCellVerticalAlign)99 });

        Assert.Contains("Table vertical alignments must be defined PDF cell vertical alignment values.", invalidVerticalAlignException.Message, StringComparison.Ordinal);

        var invalidCellAlign = PdfTableStyles.Minimal();

        var invalidCellAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCellAlign.CellAlignments = new Dictionary<(int Row, int Column), OfficeAlignment> {
                [(0, 0)] = (OfficeAlignment)99
            });

        Assert.Contains("Table column alignments must be Left, Center, Right, Start, or End.", invalidCellAlignException.Message, StringComparison.Ordinal);

        var invalidCellVerticalAlign = PdfTableStyles.Minimal();

        var invalidCellVerticalAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCellVerticalAlign.CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(0, 0)] = (PdfCellVerticalAlign)99
            });

        Assert.Contains("Table vertical alignments must be defined PDF cell vertical alignment values.", invalidCellVerticalAlignException.Message, StringComparison.Ordinal);
    }



}
