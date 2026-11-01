using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>A genuine empty source has one native CR definition, which supplies both refilled body and last-run typing.</summary>
    private static bool TryPlanEmptyRefill(PublisherBinaryData data, PublisherParseContext context,
        IReadOnlyList<RetainedParagraph> templates, OfficeRichTextParagraph after, CancellationToken token,
        PublisherQuillWriteResources? resources, out List<PublisherQuillFormattingRange> ranges, out byte[]? definition) {
        ranges = new(); definition = null;
        RetainedParagraph[] candidates = templates.Where(t => CanRefillParagraph(t.Paragraph, after)).ToArray();
        // Retained paragraph formatting gives stronger correspondence than an
        // unrelated empty paragraph whose supported layout could be rewritten.
        RetainedParagraph[] retainedFormatting = candidates.Where(t => PublisherTextComparison.ParagraphFormattingEqual(t.Paragraph, after)).ToArray();
        if (retainedFormatting.Length != 0) candidates = retainedFormatting;
        foreach (RetainedParagraph candidate in candidates) {
            token.ThrowIfCancellationRequested(); context.Record();
            var rewritten = new List<PublisherQuillFormattingRange>(); int position = 0;
            foreach (OfficeRichTextRun run in after.Runs) {
                token.ThrowIfCancellationRequested(); context.Record();
                byte[]? style = PublisherQuillCharacterStyleWriter.RewriteReplacement(data, candidate.TerminatorSpan,
                    candidate.Paragraph.Runs[0], run, context, resources);
                if (style == null) return false;
                position = checked(position + run.Text.Length * 2); AppendRange(rewritten, position, style);
            }
            // There was no source body with a distinct ending role: the source
            // CR supplies every run, including the refilled paragraph's CR.
            AppendRange(rewritten, checked(position + 2), rewritten[rewritten.Count - 1].Style);
            byte[]? paragraph = PublisherQuillParagraphStyleWriter.Rewrite(data, candidate.DefinitionSpan, candidate.Paragraph, after, context);
            if (paragraph == null || definition != null && (!definition.SequenceEqual(paragraph)
                || ranges.Count != rewritten.Count || !ranges.Zip(rewritten, (a, b) => a.AbsoluteEnd == b.AbsoluteEnd && a.Style.SequenceEqual(b.Style)).All(equal => equal))) return false;
            if (definition == null) { definition = paragraph; ranges = rewritten; }
        }
        return definition != null;
    }

    private static bool CanRefillParagraph(OfficeRichTextParagraph before, OfficeRichTextParagraph after) =>
        before.Label == null && after.Label == null && before.Runs.Count == 1 && before.Runs[0].Text.Length == 0
        && after.Runs.Count > 0 && after.Runs.All(r => r.Text.Length > 0 && PublisherTextComparison.RunPreservedFormattingEqual(before.Runs[0], r));
}
