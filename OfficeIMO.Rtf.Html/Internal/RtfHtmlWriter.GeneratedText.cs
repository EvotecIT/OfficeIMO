namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendGeneratedText(StringBuilder builder, RtfGeneratedText generatedText) {
        builder.Append("<span data-officeimo-rtf-generated-text=\"");
        builder.Append(EncodeAttribute(FormatGeneratedTextKind(generatedText.Kind)));
        builder.Append("\"></span>");
    }

    private static string FormatGeneratedTextKind(RtfGeneratedTextKind kind) {
        switch (kind) {
            case RtfGeneratedTextKind.SectionNumber:
                return "section-number";
            case RtfGeneratedTextKind.CurrentDate:
                return "current-date";
            case RtfGeneratedTextKind.CurrentDateLong:
                return "current-date-long";
            case RtfGeneratedTextKind.CurrentDateAbbreviated:
                return "current-date-abbreviated";
            case RtfGeneratedTextKind.CurrentTime:
                return "current-time";
            default:
                return "page-number";
        }
    }
}
