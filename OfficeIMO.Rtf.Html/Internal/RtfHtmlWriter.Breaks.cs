namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendBreak(StringBuilder builder, RtfBreakKind kind) {
        switch (kind) {
            case RtfBreakKind.Page:
                builder.Append("<br data-officeimo-rtf-break=\"page\" style=\"page-break-before:always;break-before:page;\">");
                break;
            case RtfBreakKind.Column:
                builder.Append("<br data-officeimo-rtf-break=\"column\" style=\"break-before:column;\">");
                break;
            default:
                builder.Append("<br>");
                break;
        }
    }
}
