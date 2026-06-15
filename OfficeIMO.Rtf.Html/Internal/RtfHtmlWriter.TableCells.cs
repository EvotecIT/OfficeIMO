namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendTableCellMetadataAttributes(StringBuilder builder, RtfTableCell cell, int inferredRightBoundaryTwips) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddDirectTableCellMetadata(values, "cell", cell, inferredRightBoundaryTwips);

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
}
