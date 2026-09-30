namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteListText(StringBuilder builder, RtfParagraph? listText, RtfWriteContext context) {
        if (listText == null || listText.Inlines.Count == 0) {
            return;
        }

        builder.Append(@"{\listtext ");
        var state = new RunWriteState(context.DefaultLanguageId);
        foreach (IRtfInline inline in listText.Inlines) {
            WriteInline(builder, inline, state, context);
        }

        builder.Append('}');
    }
}
