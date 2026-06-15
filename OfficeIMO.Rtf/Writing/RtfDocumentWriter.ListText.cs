namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteListText(StringBuilder builder, RtfParagraph? listText, int? defaultLanguageId, int unicodeSkipCount) {
        if (listText == null || listText.Inlines.Count == 0) {
            return;
        }

        builder.Append(@"{\listtext ");
        var state = new RunWriteState(defaultLanguageId);
        foreach (IRtfInline inline in listText.Inlines) {
            WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
        }

        builder.Append('}');
    }
}
