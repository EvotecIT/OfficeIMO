using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendTableCellMetadataAttributes(StringBuilder builder, RtfTableRow row, int cellIndex, int columnSpan) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddDirectTableCellMetadata(values, "cell", row.Cells[cellIndex], (cellIndex + 1) * 2400);
        AddContinuationTableCellMetadata(values, "cell.continuation", row, cellIndex, columnSpan);

        if (values.Count == 0) {
            return;
        }

        builder.Append(" data-officeimo-rtf-cell=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }

    private static void AddDirectTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell, int inferredRightBoundaryTwips) {
        if (cell.RightBoundaryTwips.HasValue && cell.RightBoundaryTwips.Value != inferredRightBoundaryTwips) {
            AddNullableInt(values, prefix + ".rightBoundary", cell.RightBoundaryTwips);
        }

        AddTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
        AddTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
    }

    private static void AddContinuationTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableRow row, int cellIndex, int columnSpan) {
        for (int offset = 1; offset < columnSpan && cellIndex + offset < row.Cells.Count; offset++) {
            AddMergedContinuationTableCellMetadata(values, prefix + "." + offset.ToString(CultureInfo.InvariantCulture), row.Cells[cellIndex + offset], (cellIndex + offset + 1) * 2400);
        }
    }

    private static void AddMergedContinuationTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell, int inferredRightBoundaryTwips) {
        if (cell.RightBoundaryTwips.HasValue && cell.RightBoundaryTwips.Value != inferredRightBoundaryTwips) {
            AddNullableInt(values, prefix + ".rightBoundary", cell.RightBoundaryTwips);
        }

        AddEnum(values, prefix + ".horizontalMerge", cell.HorizontalMerge == RtfTableCellMerge.Continue ? (RtfTableCellMerge?)null : cell.HorizontalMerge);
        AddEnum(values, prefix + ".verticalMerge", cell.VerticalMerge == RtfTableCellMerge.None ? (RtfTableCellMerge?)null : cell.VerticalMerge);
        AddNullableInt(values, prefix + ".backgroundColor", cell.BackgroundColorIndex);
        AddNullableInt(values, prefix + ".shadingForeground", cell.ShadingForegroundColorIndex);
        AddNullableInt(values, prefix + ".shadingPercent", cell.ShadingPatternPercent);
        AddEnum(values, prefix + ".shadingPattern", cell.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : cell.ShadingPattern);
        AddEnum(values, prefix + ".verticalAlignment", cell.VerticalAlignment);
        AddEnum(values, prefix + ".textFlow", cell.TextFlow);
        AddNullableInt(values, prefix + ".preferredWidth", cell.PreferredWidth);
        AddEnum(values, prefix + ".preferredWidthUnit", cell.PreferredWidthUnit);
        AddBool(values, prefix + ".hideCellMark", cell.HideCellMark);
        AddBool(values, prefix + ".noWrap", cell.NoWrap);
        AddBool(values, prefix + ".fitText", cell.FitText);
        AddNullableInt(values, prefix + ".padding.top", cell.PaddingTopTwips);
        AddNullableInt(values, prefix + ".padding.left", cell.PaddingLeftTwips);
        AddNullableInt(values, prefix + ".padding.bottom", cell.PaddingBottomTwips);
        AddNullableInt(values, prefix + ".padding.right", cell.PaddingRightTwips);
        AddTableCellBorder(values, prefix + ".border.top", cell.TopBorder);
        AddTableCellBorder(values, prefix + ".border.left", cell.LeftBorder);
        AddTableCellBorder(values, prefix + ".border.bottom", cell.BottomBorder);
        AddTableCellBorder(values, prefix + ".border.right", cell.RightBorder);
        AddTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
        AddTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
    }

    private static void AddTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell) {
        AddNullableInt(values, prefix + ".rightBoundary", cell.RightBoundaryTwips);
        AddEnum(values, prefix + ".horizontalMerge", cell.HorizontalMerge == RtfTableCellMerge.None ? (RtfTableCellMerge?)null : cell.HorizontalMerge);
        AddEnum(values, prefix + ".verticalMerge", cell.VerticalMerge == RtfTableCellMerge.None ? (RtfTableCellMerge?)null : cell.VerticalMerge);
        AddNullableInt(values, prefix + ".backgroundColor", cell.BackgroundColorIndex);
        AddNullableInt(values, prefix + ".shadingForeground", cell.ShadingForegroundColorIndex);
        AddNullableInt(values, prefix + ".shadingPercent", cell.ShadingPatternPercent);
        AddEnum(values, prefix + ".shadingPattern", cell.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : cell.ShadingPattern);
        AddEnum(values, prefix + ".verticalAlignment", cell.VerticalAlignment);
        AddEnum(values, prefix + ".textFlow", cell.TextFlow);
        AddNullableInt(values, prefix + ".preferredWidth", cell.PreferredWidth);
        AddEnum(values, prefix + ".preferredWidthUnit", cell.PreferredWidthUnit);
        AddBool(values, prefix + ".hideCellMark", cell.HideCellMark);
        AddBool(values, prefix + ".noWrap", cell.NoWrap);
        AddBool(values, prefix + ".fitText", cell.FitText);
        AddNullableInt(values, prefix + ".padding.top", cell.PaddingTopTwips);
        AddNullableInt(values, prefix + ".padding.left", cell.PaddingLeftTwips);
        AddNullableInt(values, prefix + ".padding.bottom", cell.PaddingBottomTwips);
        AddNullableInt(values, prefix + ".padding.right", cell.PaddingRightTwips);
        AddTableCellBorder(values, prefix + ".border.top", cell.TopBorder);
        AddTableCellBorder(values, prefix + ".border.left", cell.LeftBorder);
        AddTableCellBorder(values, prefix + ".border.bottom", cell.BottomBorder);
        AddTableCellBorder(values, prefix + ".border.right", cell.RightBorder);
        AddTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
        AddTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
    }
}
