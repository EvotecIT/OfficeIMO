using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AddTableStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
        if (style.Kind != RtfStyleKind.Table) {
            return;
        }

        RtfTableRow row = style.TableRowFormat;
        AddBool(values, prefix + ".repeatHeader", row.RepeatHeader);
        AddBool(values, prefix + ".keepTogether", row.KeepTogether);
        AddBool(values, prefix + ".keepWithNext", row.KeepWithNext);
        AddNullableBool(values, prefix + ".autoFit", row.AutoFit);
        AddEnum(values, prefix + ".direction", row.Direction);
        AddNullableInt(values, prefix + ".height", row.HeightTwips);
        AddNullableInt(values, prefix + ".cellGap", row.CellGapTwips);
        AddNullableInt(values, prefix + ".leftIndent", row.LeftIndentTwips);
        AddEnum(values, prefix + ".alignment", row.Alignment);
        AddNullableInt(values, prefix + ".preferredWidth", row.PreferredWidth);
        AddEnum(values, prefix + ".preferredWidthUnit", row.PreferredWidthUnit);
        AddNullableInt(values, prefix + ".backgroundColor", row.BackgroundColorIndex);
        AddNullableInt(values, prefix + ".shadingForeground", row.ShadingForegroundColorIndex);
        AddNullableInt(values, prefix + ".shadingPatternValue", row.ShadingPatternValue);
        AddNullableInt(values, prefix + ".shadingPercent", row.ShadingPatternPercent);
        AddEnum(values, prefix + ".shadingPattern", row.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : row.ShadingPattern);
        AddNullableInt(values, prefix + ".padding.top", row.PaddingTopTwips);
        AddNullableInt(values, prefix + ".padding.left", row.PaddingLeftTwips);
        AddNullableInt(values, prefix + ".padding.bottom", row.PaddingBottomTwips);
        AddNullableInt(values, prefix + ".padding.right", row.PaddingRightTwips);
        AddNullableInt(values, prefix + ".spacing.top", row.SpacingTopTwips);
        AddNullableInt(values, prefix + ".spacing.left", row.SpacingLeftTwips);
        AddNullableInt(values, prefix + ".spacing.bottom", row.SpacingBottomTwips);
        AddNullableInt(values, prefix + ".spacing.right", row.SpacingRightTwips);
        AddBool(values, prefix + ".noOverlap", row.NoOverlap);
        AddEnum(values, prefix + ".horizontalAnchor", row.HorizontalAnchor);
        AddEnum(values, prefix + ".verticalAnchor", row.VerticalAnchor);
        AddEnum(values, prefix + ".horizontalPosition", row.HorizontalPosition);
        AddNullableInt(values, prefix + ".horizontalPositionTwips", row.HorizontalPositionTwips);
        AddEnum(values, prefix + ".verticalPosition", row.VerticalPosition);
        AddNullableInt(values, prefix + ".verticalPositionTwips", row.VerticalPositionTwips);
        AddNullableInt(values, prefix + ".textWrap.left", row.TextWrapLeftTwips);
        AddNullableInt(values, prefix + ".textWrap.right", row.TextWrapRightTwips);
        AddNullableInt(values, prefix + ".textWrap.top", row.TextWrapTopTwips);
        AddNullableInt(values, prefix + ".textWrap.bottom", row.TextWrapBottomTwips);
        AddTableRowBorder(values, prefix + ".border.top", row.TopBorder);
        AddTableRowBorder(values, prefix + ".border.left", row.LeftBorder);
        AddTableRowBorder(values, prefix + ".border.bottom", row.BottomBorder);
        AddTableRowBorder(values, prefix + ".border.right", row.RightBorder);
        AddTableRowBorder(values, prefix + ".border.horizontal", row.HorizontalBorder);
        AddTableRowBorder(values, prefix + ".border.vertical", row.VerticalBorder);
        AddTableStyleCells(values, prefix + ".cell", row.Cells);
    }

    private static void AddTableStyleCells(Dictionary<string, string> values, string prefix, IReadOnlyList<RtfTableCell> cells) {
        for (int index = 0; index < cells.Count; index++) {
            RtfTableCell cell = cells[index];
            string cellPrefix = prefix + "." + index.ToString(CultureInfo.InvariantCulture);
            AddNullableInt(values, cellPrefix + ".rightBoundary", cell.RightBoundaryTwips);
            AddEnum(values, cellPrefix + ".horizontalMerge", cell.HorizontalMerge == RtfTableCellMerge.None ? (RtfTableCellMerge?)null : cell.HorizontalMerge);
            AddEnum(values, cellPrefix + ".verticalMerge", cell.VerticalMerge == RtfTableCellMerge.None ? (RtfTableCellMerge?)null : cell.VerticalMerge);
            AddNullableInt(values, cellPrefix + ".backgroundColor", cell.BackgroundColorIndex);
            AddNullableInt(values, cellPrefix + ".shadingForeground", cell.ShadingForegroundColorIndex);
            AddNullableInt(values, cellPrefix + ".shadingPercent", cell.ShadingPatternPercent);
            AddEnum(values, cellPrefix + ".shadingPattern", cell.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : cell.ShadingPattern);
            AddEnum(values, cellPrefix + ".verticalAlignment", cell.VerticalAlignment);
            AddEnum(values, cellPrefix + ".textFlow", cell.TextFlow);
            AddNullableInt(values, cellPrefix + ".preferredWidth", cell.PreferredWidth);
            AddEnum(values, cellPrefix + ".preferredWidthUnit", cell.PreferredWidthUnit);
            AddBool(values, cellPrefix + ".hideCellMark", cell.HideCellMark);
            AddBool(values, cellPrefix + ".noWrap", cell.NoWrap);
            AddBool(values, cellPrefix + ".fitText", cell.FitText);
            AddNullableInt(values, cellPrefix + ".padding.top", cell.PaddingTopTwips);
            AddNullableInt(values, cellPrefix + ".padding.left", cell.PaddingLeftTwips);
            AddNullableInt(values, cellPrefix + ".padding.bottom", cell.PaddingBottomTwips);
            AddNullableInt(values, cellPrefix + ".padding.right", cell.PaddingRightTwips);
            AddTableCellBorder(values, cellPrefix + ".border.top", cell.TopBorder);
            AddTableCellBorder(values, cellPrefix + ".border.left", cell.LeftBorder);
            AddTableCellBorder(values, cellPrefix + ".border.bottom", cell.BottomBorder);
            AddTableCellBorder(values, cellPrefix + ".border.right", cell.RightBorder);
            AddTableCellBorder(values, cellPrefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
            AddTableCellBorder(values, cellPrefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
        }
    }

    private static void AddTableRowBorder(Dictionary<string, string> values, string prefix, RtfTableRowBorder border) {
        if (!border.HasAnyValue) {
            return;
        }

        AddEnum(values, prefix + ".style", border.Style == RtfTableCellBorderStyle.None ? (RtfTableCellBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
    }

    private static void AddTableCellBorder(Dictionary<string, string> values, string prefix, RtfTableCellBorder border) {
        if (!border.HasAnyValue) {
            return;
        }

        AddEnum(values, prefix + ".style", border.Style == RtfTableCellBorderStyle.None ? (RtfTableCellBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
    }
}
