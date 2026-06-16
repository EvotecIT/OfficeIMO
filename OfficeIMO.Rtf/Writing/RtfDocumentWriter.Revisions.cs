using System.Globalization;

namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteRevisionTable(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        if (document.RevisionAuthors.Count == 0) return;

        builder.Append(@"{\*\revtbl");
        foreach (RtfRevisionAuthor author in document.RevisionAuthors) {
            builder.Append('{');
            builder.Append(EscapeText(author.Name, unicodeSkipCount));
            builder.Append(";}");
        }

        builder.Append('}');
    }

    private static void WriteRevisionSaveIdTable(StringBuilder builder, RtfDocument document) {
        if (!document.RevisionRootSaveId.HasValue && document.RevisionSaveIds.Count == 0) return;

        builder.Append(@"{\*\rsidtbl");
        if (document.RevisionRootSaveId.HasValue) {
            builder.Append(@"\rsidroot");
            builder.Append(document.RevisionRootSaveId.Value.ToString(CultureInfo.InvariantCulture));
        }

        foreach (int id in document.RevisionSaveIds) {
            builder.Append(@"\rsid");
            builder.Append(id.ToString(CultureInfo.InvariantCulture));
        }

        builder.Append('}');
    }

    private static void WriteRevisionPrefix(StringBuilder builder, RtfRun run, RunWriteState state) {
        if (state.RevisionKind != run.RevisionKind) {
            if (state.RevisionKind == RtfRevisionKind.Inserted) {
                builder.Append(@"\revised0 ");
            } else if (state.RevisionKind == RtfRevisionKind.Deleted) {
                builder.Append(@"\deleted0 ");
            }

            if (run.RevisionKind == RtfRevisionKind.Inserted) {
                builder.Append(@"\revised ");
            } else if (run.RevisionKind == RtfRevisionKind.Deleted) {
                builder.Append(@"\deleted ");
            }

            state.RevisionKind = run.RevisionKind;
        }

        if (run.RevisionAuthorIndex.HasValue && state.RevisionAuthorIndex != run.RevisionAuthorIndex.Value) {
            builder.Append(@"\revauth");
            builder.Append(run.RevisionAuthorIndex.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.RevisionAuthorIndex = run.RevisionAuthorIndex.Value;
        }

        if (run.RevisionTimestampValue.HasValue && state.RevisionTimestampValue != run.RevisionTimestampValue.Value) {
            builder.Append(@"\revdttm");
            builder.Append(run.RevisionTimestampValue.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.RevisionTimestampValue = run.RevisionTimestampValue.Value;
        }

        WriteRevisionSaveId(builder, @"\charrsid", run.CharacterRevisionSaveId, ref state.CharacterRevisionSaveId);
        WriteRevisionSaveId(builder, @"\insrsid", run.InsertionRevisionSaveId, ref state.InsertionRevisionSaveId);
        WriteRevisionSaveId(builder, @"\delrsid", run.DeletionRevisionSaveId, ref state.DeletionRevisionSaveId);
    }

    private static void WriteRevisionSaveId(StringBuilder builder, string control, int? desired, ref int? current) {
        if (desired == current) return;
        if (!desired.HasValue) {
            current = null;
            return;
        }

        builder.Append(control);
        builder.Append(desired.Value.ToString(CultureInfo.InvariantCulture));
        builder.Append(' ');
        current = desired.Value;
    }

    private static void ResetRevisionState(StringBuilder builder, RunWriteState state) {
        if (state.RevisionKind == RtfRevisionKind.Inserted) {
            builder.Append(@"\revised0 ");
        } else if (state.RevisionKind == RtfRevisionKind.Deleted) {
            builder.Append(@"\deleted0 ");
        }

        state.RevisionKind = RtfRevisionKind.None;
        state.RevisionAuthorIndex = null;
        state.RevisionTimestampValue = null;
        state.CharacterRevisionSaveId = null;
        state.InsertionRevisionSaveId = null;
        state.DeletionRevisionSaveId = null;
    }
}
