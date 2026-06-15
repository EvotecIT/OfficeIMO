using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendNote(StringBuilder builder, RtfNote? note, RtfDocument document) {
        if (note == null) {
            return;
        }

        builder.Append("<span data-officeimo-rtf-note=\"");
        builder.Append(FormatNoteKind(note.Kind));
        builder.Append("\" data-officeimo-rtf-note-content=\"");
        builder.Append(EncodeAttribute(EncodeNoteContent(note, document)));
        builder.Append('"');
        AppendNoteAttribute(builder, "data-officeimo-rtf-note-id", note.Id);
        AppendNoteAttribute(builder, "data-officeimo-rtf-note-author", note.Author);
        if (note.Created.HasValue) {
            AppendNoteAttribute(builder, "data-officeimo-rtf-note-created", note.Created.Value.ToString("O", CultureInfo.InvariantCulture));
        }

        builder.Append("></span>");
    }

    private static string EncodeNoteContent(RtfNote note, RtfDocument document) {
        var content = new StringBuilder();
        var options = new RtfHtmlSaveOptions { FragmentOnly = true };
        for (int index = 0; index < note.Paragraphs.Count; index++) {
            if (index > 0) {
                content.Append(options.GetNewLine());
            }

            AppendParagraph(content, note.Paragraphs[index], options, document);
        }

        return Convert.ToBase64String(Encoding.UTF8.GetBytes(content.ToString()));
    }

    private static string FormatNoteKind(RtfNoteKind kind) {
        switch (kind) {
            case RtfNoteKind.Annotation:
                return "annotation";
            case RtfNoteKind.Endnote:
                return "endnote";
            default:
                return "footnote";
        }
    }

    private static void AppendNoteAttribute(StringBuilder builder, string name, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(EncodeAttribute(value!));
        builder.Append('"');
    }
}
