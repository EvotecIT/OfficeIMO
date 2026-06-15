namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyTableCellMetadataAttributes(HtmlToken token) {
            Dictionary<string, string> values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-cell"));
            if (values.Count == 0 || _cell == null) {
                return;
            }

            ApplyDirectTableCellMetadata(values, "cell", _cell);
        }

        private static void ApplyDirectTableCellMetadata(Dictionary<string, string> values, string prefix, RtfTableCell cell) {
            ApplyDirectCellInt(values, prefix + ".rightBoundary", value => cell.RightBoundaryTwips = value);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
            ApplyDirectTableCellBorder(values, prefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
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
    }
}
