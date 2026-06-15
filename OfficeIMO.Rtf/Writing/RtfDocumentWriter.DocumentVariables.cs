namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteDocumentVariables(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        foreach (RtfDocumentVariable variable in document.DocumentVariables) {
            builder.Append(@"{\*\docvar {");
            builder.Append(EscapeText(variable.Name, unicodeSkipCount));
            builder.Append("}{");
            builder.Append(EscapeText(variable.Value, unicodeSkipCount));
            builder.Append("}}");
        }
    }
}
