using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendParagraphRevisionAttributes(StringBuilder builder, RtfParagraph paragraph) {
        AppendRevisionAttribute(builder, "data-officeimo-rtf-pararsid", paragraph.RevisionSaveId);
    }

    private static bool AppendRevisionStart(StringBuilder builder, RtfRun run, RtfDocument document) {
        if (!HasRevisionMetadata(run)) {
            return false;
        }

        string tagName = GetRevisionTagName(run.RevisionKind);
        builder.Append('<');
        builder.Append(tagName);
        builder.Append(" data-officeimo-rtf-revision=\"");
        builder.Append(FormatRevisionKind(run.RevisionKind));
        builder.Append('"');
        AppendRevisionAttribute(builder, "data-officeimo-rtf-revision-author-index", run.RevisionAuthorIndex);
        if (TryGetRevisionAuthor(document, run.RevisionAuthorIndex, out string? author)) {
            AppendRevisionAttribute(builder, "data-officeimo-rtf-revision-author", author);
        }

        AppendRevisionAttribute(builder, "data-officeimo-rtf-revision-timestamp", run.RevisionTimestampValue);
        AppendRevisionAttribute(builder, "data-officeimo-rtf-charrsid", run.CharacterRevisionSaveId);
        AppendRevisionAttribute(builder, "data-officeimo-rtf-insrsid", run.InsertionRevisionSaveId);
        AppendRevisionAttribute(builder, "data-officeimo-rtf-delrsid", run.DeletionRevisionSaveId);
        builder.Append('>');
        return true;
    }

    private static void AppendRevisionEnd(StringBuilder builder, RtfRun run, bool opened) {
        if (!opened) {
            return;
        }

        builder.Append("</");
        builder.Append(GetRevisionTagName(run.RevisionKind));
        builder.Append('>');
    }

    private static bool HasRevisionMetadata(RtfRun run) {
        return run.RevisionKind != RtfRevisionKind.None ||
               run.RevisionAuthorIndex.HasValue ||
               run.RevisionTimestampValue.HasValue ||
               run.CharacterRevisionSaveId.HasValue ||
               run.InsertionRevisionSaveId.HasValue ||
               run.DeletionRevisionSaveId.HasValue;
    }

    private static string GetRevisionTagName(RtfRevisionKind kind) {
        switch (kind) {
            case RtfRevisionKind.Inserted:
                return "ins";
            case RtfRevisionKind.Deleted:
                return "del";
            default:
                return "span";
        }
    }

    private static string FormatRevisionKind(RtfRevisionKind kind) {
        switch (kind) {
            case RtfRevisionKind.Inserted:
                return "inserted";
            case RtfRevisionKind.Deleted:
                return "deleted";
            default:
                return "none";
        }
    }

    private static bool TryGetRevisionAuthor(RtfDocument document, int? index, out string? author) {
        if (!index.HasValue || index.Value < 0 || index.Value >= document.RevisionAuthors.Count) {
            author = null;
            return false;
        }

        author = document.RevisionAuthors[index.Value].Name;
        return !string.IsNullOrEmpty(author);
    }

    private static void AppendRevisionAttribute(StringBuilder builder, string name, int? value) {
        if (!value.HasValue) {
            return;
        }

        AppendRevisionAttribute(builder, name, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static void AppendRevisionAttribute(StringBuilder builder, string name, string? value) {
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
