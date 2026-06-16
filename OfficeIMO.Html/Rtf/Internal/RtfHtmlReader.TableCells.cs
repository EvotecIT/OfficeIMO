using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyTableCellMetadataAttributes(IElement token, int columnStart, int columnSpan) {
            Dictionary<string, string> values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-cell"));
            if (values.Count == 0 || _cell == null) {
                return;
            }

            ApplyDirectTableCellMetadata(values, "cell", _cell);
            ApplyContinuationTableCellMetadata(values, "cell.continuation", columnStart, columnSpan);
        }

        private static void ApplyDirectTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell) {
            ApplyDirectCellInt(values, prefix + ".rightBoundary", value => cell.RightBoundaryTwips = value);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
        }

        private void ApplyContinuationTableCellMetadata(Dictionary<string, string> values, string prefix, int columnStart, int columnSpan) {
            if (_row == null) {
                return;
            }

            for (int offset = 1; offset < columnSpan && columnStart + offset < _row.Cells.Count; offset++) {
                string continuationPrefix = prefix + "." + offset.ToString(CultureInfo.InvariantCulture);
                if (HasTableCellMetadata(values, continuationPrefix)) {
                    ApplyMergedContinuationTableCellMetadata(values, continuationPrefix, _row.Cells[columnStart + offset]);
                }
            }
        }

        private static void ApplyMergedContinuationTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell) {
            ApplyDirectCellInt(values, prefix + ".rightBoundary", value => cell.RightBoundaryTwips = value);
            ApplyDirectCellEnum<RtfTableCellMerge>(values, prefix + ".horizontalMerge", value => cell.HorizontalMerge = value.GetValueOrDefault());
            ApplyDirectCellEnum<RtfTableCellMerge>(values, prefix + ".verticalMerge", value => cell.VerticalMerge = value.GetValueOrDefault());
            ApplyDirectCellInt(values, prefix + ".backgroundColor", value => cell.BackgroundColorIndex = value);
            ApplyDirectCellInt(values, prefix + ".shadingForeground", value => cell.ShadingForegroundColorIndex = value);
            ApplyDirectCellInt(values, prefix + ".shadingPercent", value => cell.ShadingPatternPercent = value);
            ApplyDirectCellEnum<RtfShadingPattern>(values, prefix + ".shadingPattern", value => cell.ShadingPattern = value.GetValueOrDefault());
            ApplyDirectCellEnum<RtfTableCellVerticalAlignment>(values, prefix + ".verticalAlignment", value => cell.VerticalAlignment = value);
            ApplyDirectCellEnum<RtfTableCellTextFlow>(values, prefix + ".textFlow", value => cell.TextFlow = value);
            ApplyDirectCellInt(values, prefix + ".preferredWidth", value => cell.PreferredWidth = value);
            ApplyDirectCellEnum<RtfTableWidthUnit>(values, prefix + ".preferredWidthUnit", value => cell.PreferredWidthUnit = value);
            ApplyDirectCellBool(values, prefix + ".hideCellMark", value => cell.HideCellMark = value);
            ApplyDirectCellBool(values, prefix + ".noWrap", value => cell.NoWrap = value);
            ApplyDirectCellBool(values, prefix + ".fitText", value => cell.FitText = value);
            ApplyDirectCellInt(values, prefix + ".padding.top", value => cell.PaddingTopTwips = value);
            ApplyDirectCellInt(values, prefix + ".padding.left", value => cell.PaddingLeftTwips = value);
            ApplyDirectCellInt(values, prefix + ".padding.bottom", value => cell.PaddingBottomTwips = value);
            ApplyDirectCellInt(values, prefix + ".padding.right", value => cell.PaddingRightTwips = value);
            ApplyDirectTableCellBorder(values, prefix + ".border.top", cell.TopBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.left", cell.LeftBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.bottom", cell.BottomBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.right", cell.RightBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
        }

        private static void ApplyTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell) {
            cell.RightBoundaryTwips = ReadInt(values, prefix + ".rightBoundary");
            cell.HorizontalMerge = ReadEnum(values, prefix + ".horizontalMerge", RtfTableCellMerge.None);
            cell.VerticalMerge = ReadEnum(values, prefix + ".verticalMerge", RtfTableCellMerge.None);
            cell.BackgroundColorIndex = ReadInt(values, prefix + ".backgroundColor");
            cell.ShadingForegroundColorIndex = ReadInt(values, prefix + ".shadingForeground");
            cell.ShadingPatternPercent = ReadInt(values, prefix + ".shadingPercent");
            cell.ShadingPattern = ReadEnum(values, prefix + ".shadingPattern", RtfShadingPattern.None);
            cell.VerticalAlignment = ReadEnum<RtfTableCellVerticalAlignment>(values, prefix + ".verticalAlignment");
            cell.TextFlow = ReadEnum<RtfTableCellTextFlow>(values, prefix + ".textFlow");
            cell.PreferredWidth = ReadInt(values, prefix + ".preferredWidth");
            cell.PreferredWidthUnit = ReadEnum<RtfTableWidthUnit>(values, prefix + ".preferredWidthUnit");
            cell.HideCellMark = ReadBool(values, prefix + ".hideCellMark") == true;
            cell.NoWrap = ReadBool(values, prefix + ".noWrap") == true;
            cell.FitText = ReadBool(values, prefix + ".fitText") == true;
            cell.PaddingTopTwips = ReadInt(values, prefix + ".padding.top");
            cell.PaddingLeftTwips = ReadInt(values, prefix + ".padding.left");
            cell.PaddingBottomTwips = ReadInt(values, prefix + ".padding.bottom");
            cell.PaddingRightTwips = ReadInt(values, prefix + ".padding.right");
            ApplyTableCellBorder(values, prefix + ".border.top", cell.TopBorder);
            ApplyTableCellBorder(values, prefix + ".border.left", cell.LeftBorder);
            ApplyTableCellBorder(values, prefix + ".border.bottom", cell.BottomBorder);
            ApplyTableCellBorder(values, prefix + ".border.right", cell.RightBorder);
            ApplyTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
            ApplyTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
        }

        private static void ApplyDirectTableCellBorder(Dictionary<string, string> values, string prefix, RtfTableCellBorder border) {
            if (!values.Keys.Any(key => key.StartsWith(prefix + ".", StringComparison.Ordinal))) {
                return;
            }

            ApplyTableCellBorder(values, prefix, border);
        }

        private static void ApplyDirectCellInt(Dictionary<string, string> values, string key, Action<int?> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadInt(values, key));
            }
        }

        private static void ApplyDirectCellBool(Dictionary<string, string> values, string key, Action<bool> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadBool(values, key) == true);
            }
        }

        private static void ApplyDirectCellEnum<T>(Dictionary<string, string> values, string key, Action<T?> assign) where T : struct {
            if (values.ContainsKey(key)) {
                assign(ReadEnum<T>(values, key));
            }
        }

        private static bool HasTableCellMetadata(Dictionary<string, string> values, string prefix) {
            return values.Keys.Any(key => key.StartsWith(prefix + ".", StringComparison.Ordinal));
        }
    }
}
