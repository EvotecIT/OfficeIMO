namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendBreak(StringBuilder builder, RtfBreakKind kind, bool includeRoundTripMetadata) {
        switch (kind) {
            case RtfBreakKind.SoftLine:
                builder.Append(includeRoundTripMetadata ? "<br data-officeimo-rtf-break=\"soft-line\">" : "<br>");
                break;
            case RtfBreakKind.Page:
                builder.Append(includeRoundTripMetadata ? "<br data-officeimo-rtf-break=\"page\" style=\"page-break-before:always;break-before:page;\">" : "<br style=\"page-break-before:always;break-before:page;\">");
                break;
            case RtfBreakKind.SoftPage:
                builder.Append(includeRoundTripMetadata ? "<br data-officeimo-rtf-break=\"soft-page\">" : "<br>");
                break;
            case RtfBreakKind.Column:
                builder.Append(includeRoundTripMetadata ? "<br data-officeimo-rtf-break=\"column\" style=\"break-before:column;\">" : "<br style=\"break-before:column;\">");
                break;
            default:
                builder.Append("<br>");
                break;
        }
    }
}
