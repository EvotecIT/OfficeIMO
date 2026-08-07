namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendBookmarkMarker(StringBuilder builder, RtfBookmarkMarker marker, bool includeRoundTripMetadata) {
        if (marker.Kind == RtfBookmarkMarkerKind.Start) {
            builder.Append("<a id=\"");
            builder.Append(EncodeAttribute(marker.Name));
            if (includeRoundTripMetadata) {
                builder.Append("\" data-officeimo-rtf-bookmark=\"start\" data-officeimo-rtf-bookmark-name=\"");
                builder.Append(EncodeAttribute(marker.Name));
            }
            builder.Append("\"></a>");
            return;
        }

        if (!includeRoundTripMetadata) {
            return;
        }

        builder.Append("<a data-officeimo-rtf-bookmark=\"end\" data-officeimo-rtf-bookmark-name=\"");
        builder.Append(EncodeAttribute(marker.Name));
        builder.Append("\"></a>");
    }
}
