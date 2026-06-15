namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteUserProperties(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        if (document.UserProperties.Count == 0) return;

        builder.Append(@"{\*\userprops");
        foreach (RtfUserProperty property in document.UserProperties) {
            builder.Append(@"{\propname ");
            builder.Append(EscapeText(property.Name, unicodeSkipCount));
            builder.Append('}');
            AppendOptionalTwips(builder, @"\proptype", property.TypeCode);
            WriteUserPropertyValue(builder, "staticval", property.StaticValue, unicodeSkipCount);
            WriteUserPropertyValue(builder, "linkval", property.LinkedValue, unicodeSkipCount);
        }

        builder.Append('}');
    }

    private static void WriteUserPropertyValue(StringBuilder builder, string destination, string? value, int unicodeSkipCount) {
        if (string.IsNullOrEmpty(value)) return;
        builder.Append(@"{\");
        builder.Append(destination);
        builder.Append(' ');
        builder.Append(EscapeText(value!, unicodeSkipCount));
        builder.Append('}');
    }
}
