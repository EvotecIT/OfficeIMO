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
    public void TableStyle_AlignmentListsSnapshotAssignedCollections() {
        var horizontal = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Center };
        var vertical = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Top, PdfCellVerticalAlign.Middle };
        var cellHorizontal = new Dictionary<(int Row, int Column), PdfColumnAlign> {
            [(1, 1)] = PdfColumnAlign.Right
        };
        var cellVertical = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
            [(1, 1)] = PdfCellVerticalAlign.Bottom
        };
        var style = TableStyles.Minimal();

        style.Alignments = horizontal;
        style.VerticalAlignments = vertical;
        style.CellAlignments = cellHorizontal;
        style.CellVerticalAlignments = cellVertical;

        horizontal[0] = PdfColumnAlign.Right;
        vertical[0] = PdfCellVerticalAlign.Bottom;
        cellHorizontal[(1, 1)] = PdfColumnAlign.Center;
        cellVertical[(1, 1)] = PdfCellVerticalAlign.Middle;

        Assert.NotNull(style.Alignments);
        Assert.NotNull(style.VerticalAlignments);
        Assert.NotNull(style.CellAlignments);
        Assert.NotNull(style.CellVerticalAlignments);
        Assert.Equal(PdfColumnAlign.Left, style.Alignments![0]);
        Assert.Equal(PdfCellVerticalAlign.Top, style.VerticalAlignments![0]);
        Assert.Equal(PdfColumnAlign.Right, style.CellAlignments![(1, 1)]);
        Assert.Equal(PdfCellVerticalAlign.Bottom, style.CellVerticalAlignments![(1, 1)]);
    }

    [Fact]
    public void TableStyle_ColumnSizingListsSnapshotAssignedCollections() {
        var fixedWidths = new List<double?> { 60, null };
        var minWidths = new List<double?> { 40, null };
        var maxWidths = new List<double?> { null, 120 };
        var weights = new List<double> { 1, 2 };
        var rowMinHeights = new List<double?> { 18, null, 36 };
        var rowBreakPolicies = new List<bool?> { false, null, true };
        var style = TableStyles.Minimal();

        style.ColumnWidthPoints = fixedWidths;
        style.ColumnMinWidthPoints = minWidths;
        style.ColumnMaxWidthPoints = maxWidths;
        style.ColumnWidthWeights = weights;
        style.RowMinHeights = rowMinHeights;
        style.RowAllowBreakAcrossPages = rowBreakPolicies;

        fixedWidths[0] = 10;
        minWidths[0] = 10;
        maxWidths[1] = 10;
        weights[1] = 10;
        rowMinHeights[0] = 99;
        rowBreakPolicies[0] = true;

        Assert.NotNull(style.ColumnWidthPoints);
        Assert.NotNull(style.ColumnMinWidthPoints);
        Assert.NotNull(style.ColumnMaxWidthPoints);
        Assert.NotNull(style.ColumnWidthWeights);
        Assert.NotNull(style.RowMinHeights);
        Assert.NotNull(style.RowAllowBreakAcrossPages);
        Assert.Equal(60, style.ColumnWidthPoints![0]);
        Assert.Equal(40, style.ColumnMinWidthPoints![0]);
        Assert.Equal(120, style.ColumnMaxWidthPoints![1]);
        Assert.Equal(2, style.ColumnWidthWeights![1]);
        Assert.Equal(18, style.RowMinHeights![0]);
        Assert.Null(style.RowMinHeights![1]);
        Assert.Equal(36, style.RowMinHeights![2]);
        Assert.False(style.RowAllowBreakAcrossPages![0]);
        Assert.Null(style.RowAllowBreakAcrossPages![1]);
        Assert.True(style.RowAllowBreakAcrossPages![2]);
    }

    [Fact]
    public void TableStyle_FillAndBorderCollectionsSnapshotAssignedValues() {
        var columnFills = new List<PdfColor?> { PdfColor.Gray, PdfColor.LightGray };
        var cellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 1)] = new PdfColor(0.1, 0.2, 0.3)
        };
        var cellDataBar = new PdfCellDataBar {
            Color = new PdfColor(0.2, 0.3, 0.4),
            Ratio = 0.25
        };
        var cellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(1, 1)] = cellDataBar
        };
        var cellIcon = new PdfCellIcon {
            Kind = PdfCellIconKind.Circle,
            Color = new PdfColor(0.25, 0.35, 0.45),
            Size = 9,
            OffsetX = 1.25,
            OffsetY = -0.5
        };
        var cellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(1, 1)] = cellIcon
        };
        var cellBorder = new PdfCellBorder {
            Color = new PdfColor(0.4, 0.5, 0.6),
            Width = 1.25,
            DashStyle = OfficeStrokeDashStyle.Dash,
            LineStyle = PdfCellBorderLineStyle.TwoLine,
            Left = false,
            DiagonalUp = true,
            DiagonalUpBorder = new PdfCellBorderSide {
                Color = new PdfColor(0.11, 0.22, 0.33),
                Width = 1.75,
                LineStyle = PdfCellBorderLineStyle.TwoLine
            },
            TopBorder = new PdfCellBorderSide {
                Color = new PdfColor(0.7, 0.8, 0.9),
                Width = 2.25,
                DashStyle = OfficeStrokeDashStyle.Dot,
                LineStyle = PdfCellBorderLineStyle.TwoLine
            }
        };
        var cellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 1)] = cellBorder
        };
        var cellPadding = new PdfCellPadding {
            Left = 7,
            Right = 8,
            Top = 9,
            Bottom = 10
        };
        var cellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
            [(1, 1)] = cellPadding
        };
        var style = TableStyles.Minimal();

        style.BodyColumnFills = columnFills;
        style.CellFills = cellFills;
        style.CellDataBars = cellDataBars;
        style.CellIcons = cellIcons;
        style.CellBorders = cellBorders;
        style.CellPaddings = cellPaddings;

        columnFills[0] = PdfColor.White;
        cellFills[(1, 1)] = PdfColor.Black;
        cellDataBar.Ratio = 0.75;
        cellDataBar.Color = PdfColor.Black;
        cellIcon.Kind = PdfCellIconKind.Diamond;
        cellIcon.Color = PdfColor.Black;
        cellIcon.Size = 12;
        cellIcon.OffsetX = 3;
        cellIcon.OffsetY = 4;
        cellBorder.Width = 4;
        cellBorder.LineStyle = PdfCellBorderLineStyle.Standard;
        cellBorder.Left = true;
        cellBorder.DiagonalUp = false;
        cellBorder.TopBorder = new PdfCellBorderSide {
            Color = PdfColor.Black,
            Width = 5
        };
        cellPadding.Left = 30;
        cellPadding.Top = 31;

        Assert.NotNull(style.BodyColumnFills);
        Assert.NotNull(style.CellFills);
        Assert.NotNull(style.CellDataBars);
        Assert.NotNull(style.CellIcons);
        Assert.NotNull(style.CellBorders);
        Assert.NotNull(style.CellPaddings);
        Assert.Equal(PdfColor.Gray, style.BodyColumnFills![0]);
        Assert.Equal(new PdfColor(0.1, 0.2, 0.3), style.CellFills![(1, 1)]);
        Assert.Equal(new PdfColor(0.2, 0.3, 0.4), style.CellDataBars![(1, 1)].Color);
        Assert.Equal(0.25, style.CellDataBars![(1, 1)].Ratio);
        Assert.Equal(PdfCellIconKind.Circle, style.CellIcons![(1, 1)].Kind);
        Assert.Equal(new PdfColor(0.25, 0.35, 0.45), style.CellIcons![(1, 1)].Color);
        Assert.Equal(9, style.CellIcons![(1, 1)].Size);
        Assert.Equal(1.25, style.CellIcons![(1, 1)].OffsetX);
        Assert.Equal(-0.5, style.CellIcons![(1, 1)].OffsetY);
        Assert.Equal(1.25, style.CellBorders![(1, 1)].Width);
        Assert.Equal(OfficeStrokeDashStyle.Dash, style.CellBorders![(1, 1)].DashStyle);
        Assert.Equal(PdfCellBorderLineStyle.TwoLine, style.CellBorders![(1, 1)].LineStyle);
        Assert.False(style.CellBorders![(1, 1)].Left);
        Assert.True(style.CellBorders![(1, 1)].DiagonalUp);
        Assert.Equal(new PdfColor(0.11, 0.22, 0.33), style.CellBorders![(1, 1)].DiagonalUpBorder!.Color);
        Assert.Equal(1.75, style.CellBorders![(1, 1)].DiagonalUpBorder!.Width);
        Assert.Equal(PdfCellBorderLineStyle.TwoLine, style.CellBorders![(1, 1)].DiagonalUpBorder!.LineStyle);
        Assert.Equal(new PdfColor(0.7, 0.8, 0.9), style.CellBorders![(1, 1)].TopBorder!.Color);
        Assert.Equal(2.25, style.CellBorders![(1, 1)].TopBorder!.Width);
        Assert.Equal(OfficeStrokeDashStyle.Dot, style.CellBorders![(1, 1)].TopBorder!.DashStyle);
        Assert.Equal(PdfCellBorderLineStyle.TwoLine, style.CellBorders![(1, 1)].TopBorder!.LineStyle);
        Assert.Equal(7, style.CellPaddings![(1, 1)].Left);
        Assert.Equal(8, style.CellPaddings![(1, 1)].Right);
        Assert.Equal(9, style.CellPaddings![(1, 1)].Top);
        Assert.Equal(10, style.CellPaddings![(1, 1)].Bottom);
    }

    [Fact]
    public void TableStyle_TypographySettingsSurviveClone() {
        var style = TableStyles.Minimal();
        style.FontSize = 8;
        style.LineHeight = 1.6;
        style.HeaderFontSize = 12;
        style.FooterFontSize = 10;
        style.HeaderBold = false;
        style.FooterBold = false;
        style.KeepTogether = true;
        style.KeepWithNext = true;
        style.AllowRowBreakAcrossPages = false;
        style.MaxWidth = 180;
        style.LeftIndent = 24;
        style.CellPaddingLeft = 7;
        style.CellPaddingRight = 8;
        style.CellPaddingTop = 9;
        style.CellPaddingBottom = 10;
        style.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
            [(0, 0)] = new PdfCellPadding { Left = 12, Top = 13 }
        };
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(0, 0)] = new PdfCellDataBar {
                Color = new PdfColor(0.2, 0.3, 0.4),
                Ratio = 0.5
            }
        };
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(0, 0)] = new PdfCellIcon {
                Kind = PdfCellIconKind.Diamond,
                Color = new PdfColor(0.2, 0.3, 0.4),
                Size = 9,
                OffsetX = 1.25,
                OffsetY = -0.5
            }
        };
        style.CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
            [(0, 0)] = PdfColumnAlign.Right
        };
        style.CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
            [(0, 0)] = PdfCellVerticalAlign.Bottom
        };
        style.CellSpacing = 11;
        style.PreferredWidth = 160;
        style.RowMinHeights = new List<double?> { 16, null, 48 };
        style.RowAllowBreakAcrossPages = new List<bool?> { false, null, true };

        PdfTableStyle clone = style.Clone();

        Assert.Equal(8, clone.FontSize);
        Assert.Equal(1.6, clone.LineHeight);
        Assert.Equal(12, clone.HeaderFontSize);
        Assert.Equal(10, clone.FooterFontSize);
        Assert.False(clone.HeaderBold);
        Assert.False(clone.FooterBold);
        Assert.True(clone.KeepTogether);
        Assert.True(clone.KeepWithNext);
        Assert.False(clone.AllowRowBreakAcrossPages);
        Assert.Equal(160, clone.PreferredWidth);
        Assert.Equal(180, clone.MaxWidth);
        Assert.Equal(24, clone.LeftIndent);
        Assert.Equal(7, clone.CellPaddingLeft);
        Assert.Equal(8, clone.CellPaddingRight);
        Assert.Equal(9, clone.CellPaddingTop);
        Assert.Equal(10, clone.CellPaddingBottom);
        Assert.NotNull(clone.CellPaddings);
        Assert.Equal(12, clone.CellPaddings![(0, 0)].Left);
        Assert.Equal(13, clone.CellPaddings![(0, 0)].Top);
        Assert.NotNull(clone.CellDataBars);
        Assert.Equal(new PdfColor(0.2, 0.3, 0.4), clone.CellDataBars![(0, 0)].Color);
        Assert.Equal(0.5, clone.CellDataBars![(0, 0)].Ratio);
        Assert.NotNull(clone.CellIcons);
        Assert.Equal(PdfCellIconKind.Diamond, clone.CellIcons![(0, 0)].Kind);
        Assert.Equal(new PdfColor(0.2, 0.3, 0.4), clone.CellIcons![(0, 0)].Color);
        Assert.Equal(9, clone.CellIcons![(0, 0)].Size);
        Assert.Equal(1.25, clone.CellIcons![(0, 0)].OffsetX);
        Assert.Equal(-0.5, clone.CellIcons![(0, 0)].OffsetY);
        Assert.NotNull(clone.CellAlignments);
        Assert.NotNull(clone.CellVerticalAlignments);
        Assert.Equal(PdfColumnAlign.Right, clone.CellAlignments![(0, 0)]);
        Assert.Equal(PdfCellVerticalAlign.Bottom, clone.CellVerticalAlignments![(0, 0)]);
        Assert.Equal(11, clone.CellSpacing);
        Assert.NotNull(clone.RowMinHeights);
        Assert.Equal(16, clone.RowMinHeights![0]);
        Assert.Null(clone.RowMinHeights![1]);
        Assert.Equal(48, clone.RowMinHeights![2]);
        Assert.NotNull(clone.RowAllowBreakAcrossPages);
        Assert.False(clone.RowAllowBreakAcrossPages![0]);
        Assert.Null(clone.RowAllowBreakAcrossPages![1]);
        Assert.True(clone.RowAllowBreakAcrossPages![2]);
    }


}
