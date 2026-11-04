using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>New text can replace non-positional native bodies only when every rewritten source definition agrees.</summary>
    private static bool TryPlanReplacementRanges(PublisherBinaryData data, PublisherParseContext context,
        IReadOnlyList<PublisherQuillStyleUse> characters, int nativeStart, OfficeRichTextParagraph before, OfficeRichTextParagraph after,
        CancellationToken token, PublisherQuillWriteResources? resources, out List<PublisherQuillFormattingRange> ranges) {
        ranges = new();
        if (!CanReplaceText(before, after)) return false;
        int length = before.Runs.Sum(r => r.Text.Length), end = checked(nativeStart + length), cursor = nativeStart;
        var definitions = new List<PublisherQuillSourceSpan>();
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        foreach (PublisherQuillStyleUse use in characters) {
            token.ThrowIfCancellationRequested();
            if ((long)use.NativeStart + use.NativeLength <= cursor) continue;
            if (cursor == end) break;
            if (use.NativeStart > cursor) return false;
            int next = (int)Math.Min(end, (long)use.NativeStart + use.NativeLength);
            if (next <= cursor) return false;
            context.Record();
            if (distinct.Add(Convert.ToBase64String(data.Bytes, use.DirectDefinition.Offset, use.DirectDefinition.Length)))
                definitions.Add(use.DirectDefinition);
            cursor = next;
        }
        if (cursor != end || definitions.Count == 0) return false;
        int position = 0;
        foreach (OfficeRichTextRun run in after.Runs) {
            byte[]? agreed = null;
            foreach (PublisherQuillSourceSpan definition in definitions) {
                token.ThrowIfCancellationRequested();
                byte[]? rewritten = PublisherQuillCharacterStyleWriter.RewriteReplacement(data, definition, before.Runs[0], run, context, resources);
                if (rewritten == null || agreed != null && !agreed.SequenceEqual(rewritten)) return false;
                agreed = rewritten;
            }
            position = checked(position + run.Text.Length * 2);
            AppendRange(ranges, position, agreed!);
        }
        return true;
    }

    private static bool CanReplaceText(OfficeRichTextParagraph before, OfficeRichTextParagraph after) =>
        before.Label == null && after.Label == null && before.Runs.Count > 0 && after.Runs.Count > 0
        && before.Runs.All(r => r.Text.Length > 0 && PublisherTextComparison.RunPreservedFormattingEqual(before.Runs[0], r))
        && after.Runs.All(r => r.Text.Length > 0 && PublisherTextComparison.RunPreservedFormattingEqual(before.Runs[0], r))
        && string.Concat(before.Runs.Select(r => r.Text)) != string.Concat(after.Runs.Select(r => r.Text));
}
