using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Reformats intersections of source runs, requested runs and native PC ranges without moving opaque metadata.</summary>
    private static bool TryPlanContentRanges(PublisherBinaryData data, PublisherParseContext context,
        IReadOnlyList<PublisherQuillStyleUse> characters, int nativeStart, OfficeRichTextParagraph before, OfficeRichTextParagraph after,
        CancellationToken token, PublisherQuillWriteResources? resources, out List<PublisherQuillFormattingRange> ranges) {
        ranges = new();
        if (!PublisherTextComparison.ParagraphContent.Equals(before, after) || before.Runs.Count == 0
            || before.Runs.Any(r => r.Text.Length == 0) || after.Runs.Any(r => r.Text.Length == 0)) return false;
        int length = before.Runs.Sum(r => r.Text.Length), position = 0, source = 0, target = 0, native = 0;
        int sourceEnd = before.Runs[0].Text.Length, targetEnd = after.Runs[0].Text.Length;
        while (position < length) {
            token.ThrowIfCancellationRequested();
            int absolute = checked(nativeStart + position);
            while (native < characters.Count && (long)characters[native].NativeStart + characters[native].NativeLength <= absolute) native++;
            if (native == characters.Count || characters[native].NativeStart > absolute) return false;
            PublisherQuillStyleUse use = characters[native];
            int nativeEnd = (int)Math.Min(length, (long)use.NativeStart + use.NativeLength - nativeStart);
            int end = Math.Min(nativeEnd, Math.Min(sourceEnd, targetEnd));
            if (end <= position) return false;
            byte[]? style = PublisherQuillCharacterStyleWriter.Rewrite(data, use.DirectDefinition, before.Runs[source], after.Runs[target], context, resources);
            if (style == null) return false;
            AppendRange(ranges, checked(end * 2), style);
            position = end;
            if (position == sourceEnd && ++source < before.Runs.Count) sourceEnd = checked(sourceEnd + before.Runs[source].Text.Length);
            if (position == targetEnd && ++target < after.Runs.Count) targetEnd = checked(targetEnd + after.Runs[target].Text.Length);
        }
        return position == length && source == before.Runs.Count && target == after.Runs.Count;
    }
}
