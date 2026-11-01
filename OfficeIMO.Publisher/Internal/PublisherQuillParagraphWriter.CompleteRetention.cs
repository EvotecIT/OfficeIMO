using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Rewrites complete nonempty bodies while retaining the established independent interior/final CR roles.</summary>
    private static bool TryPlanCompleteRetention(PublisherBinaryData data, PublisherParseContext context,
        IReadOnlyList<RetainedParagraph> templates, IReadOnlyList<PublisherQuillStyleUse> characters,
        OfficeRichTextParagraph after, byte[] final, bool last, CancellationToken token,
        PublisherQuillWriteResources? resources, out List<PublisherQuillFormattingRange> ranges, out byte[]? definition) {
        ranges = new(); definition = null;
        RetainedParagraph[] candidates = templates.Where(t => CanReplaceText(t.Paragraph, after)).ToArray();
        RetainedParagraph[] retainedFormatting = candidates.Where(t => PublisherTextComparison.ParagraphFormattingEqual(t.Paragraph, after)).ToArray();
        if (retainedFormatting.Length != 0) candidates = retainedFormatting;
        foreach (RetainedParagraph candidate in candidates) {
            token.ThrowIfCancellationRequested(); context.Record();
            if (!TryPlanReplacementRanges(data, context, characters, candidate.NativeStart, candidate.Paragraph, after, token, resources, out var rewritten)) return false;
            byte[]? ending = last ? final : candidate.InteriorTerminator;
            if (ending == null) return false;
            // Visible body styling does not replace nonempty CR typing. Apply
            // only the nullable-attribute changes used by retained templates.
            ending = PublisherQuillCharacterStyleWriter.RewriteExplicitAttributes(ending,
                candidate.Paragraph.Runs[candidate.Paragraph.Runs.Count - 1], after.Runs[after.Runs.Count - 1], context);
            byte[]? paragraph = PublisherQuillParagraphStyleWriter.Rewrite(data, candidate.DefinitionSpan, candidate.Paragraph, after, context);
            if (ending == null || paragraph == null) return false;
            AppendRange(rewritten, checked(after.Runs.Sum(r => r.Text.Length) * 2 + 2), ending);
            if (definition != null && (!definition.SequenceEqual(paragraph) || ranges.Count != rewritten.Count
                || !ranges.Zip(rewritten, (a, b) => a.AbsoluteEnd == b.AbsoluteEnd && a.Style.SequenceEqual(b.Style)).All(equal => equal))) return false;
            if (definition == null) { definition = paragraph; ranges = rewritten; }
        }
        return definition != null;
    }
}
