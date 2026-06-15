using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AddTableStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
        if (style.Kind != RtfStyleKind.Table) {
            return;
        }

        RtfTableRow row = style.TableRowFormat;
        AddTableRowMetadata(values, prefix, row);
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
