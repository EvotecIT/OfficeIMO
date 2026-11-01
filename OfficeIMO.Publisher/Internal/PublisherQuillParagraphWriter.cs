using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Plans source-bound paragraph/character ranges and uniform-source topology changes, retaining native OPLs and CR formatting.</summary>
internal static partial class PublisherQuillParagraphWriter {
    internal static bool TryWrite(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int textOffset, CancellationToken token,
        out byte[]? characterPage, out byte[]? paragraphPage) {
        characterPage = paragraphPage = null;
        if (!TryPlan(data, context, binding, before, after, textOffset, token, out var pc, out var pp)) return false;
        characterPage = Pack(pc); paragraphPage = Pack(pp);
        return characterPage != null && paragraphPage != null;
    }

    internal static bool TryPlan(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int outputTextOffset, CancellationToken token,
        out List<PublisherQuillFormattingRange> pc, out List<PublisherQuillFormattingRange> pp, PublisherQuillWriteResources? resources = null,
        PublisherQuillStyles? typingStyles = null) {
        pc = new(); pp = new();
        if (before.Paragraphs.Count != after.Paragraphs.Count) {
            if (TryPlanTopology(data, context, binding, before, after, outputTextOffset, token, pc, pp, resources, typingStyles)) return true;
            pc.Clear(); pp.Clear();
            return TryPlanRetention(data, context, binding, before, after, outputTextOffset, token, pc, pp, resources);
        }
        PublisherQuillStyleUse[] characters = binding.Uses.Where(use => use.Tag == "FDPC").ToArray();
        PublisherQuillStyleUse[] paragraphs = binding.Uses.Where(use => use.Tag == "FDPP").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count) return false;
        int originalPosition = binding.NativeStart, position = outputTextOffset;
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++) {
            token.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeRichTextRun> oldRuns = before.Paragraphs[paragraphIndex].Runs;
            IReadOnlyList<OfficeRichTextRun> newRuns = after.Paragraphs[paragraphIndex].Runs;
            int oldLength = oldRuns.Sum(run => run.Text.Length);
            PublisherQuillStyleUse paragraph = paragraphs[paragraphIndex];
            if (paragraph.NativeStart != originalPosition || paragraph.NativeLength != oldLength + 1) return false;
            OfficeRichTextParagraph oldParagraph = before.Paragraphs[paragraphIndex], newParagraph = after.Paragraphs[paragraphIndex];
            bool resegment = !PublisherTextComparison.ParagraphTexts.Equals(oldParagraph, newParagraph)
                && PublisherTextComparison.ParagraphContent.Equals(oldParagraph, newParagraph);
            if (oldRuns.Count > 1 && oldRuns.Count != newRuns.Count && CanReplaceText(oldParagraph, newParagraph)) {
                if (!TryPlanReplacementRanges(data, context, characters, originalPosition, oldParagraph, newParagraph, token, resources, out var replacement)) return false;
                PublisherQuillStyleUse? originalCr = FindCharacter(checked(originalPosition + oldLength), 1);
                if (originalCr == null) return false;
                foreach (var range in replacement) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
                position = checked(position + newRuns.Sum(r => r.Text.Length) * 2 + 2);
                byte[]? ending = PublisherQuillCharacterStyleWriter.Rewrite(data, originalCr.Value.DirectDefinition,
                    oldRuns[oldRuns.Count - 1], newRuns[newRuns.Count - 1], context, resources);
                byte[]? paragraphStyle = PublisherQuillParagraphStyleWriter.Rewrite(data, paragraph.DirectDefinition, oldParagraph, newParagraph, context);
                if (ending == null || paragraphStyle == null) return false;
                AppendRange(pc, position, ending);
                pp.Add(new PublisherQuillFormattingRange(position, paragraphStyle));
                originalPosition = checked(originalPosition + oldLength + 1);
                continue;
            }
            // Preserve the established uniform single-run path. New interval
            // mapping covers merges and native boundaries that it cannot map.
            if (resegment && (oldRuns.Count > 1 || FindCharacter(originalPosition, Math.Max(1, oldLength)) == null)) {
                if (!TryPlanContentRanges(data, context, characters, originalPosition, oldParagraph, newParagraph, token, resources, out var mapped)) return false;
                PublisherQuillStyleUse? originalCr = FindCharacter(checked(originalPosition + oldLength), 1);
                if (originalCr == null) return false;
                foreach (var range in mapped) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
                position = checked(position + (oldLength + 1) * 2);
                AppendRange(pc, position, Copy(originalCr.Value.DirectDefinition));
                byte[]? rewrittenParagraph = PublisherQuillParagraphStyleWriter.Rewrite(data, paragraph.DirectDefinition, oldParagraph, newParagraph, context);
                if (rewrittenParagraph == null) return false;
                pp.Add(new PublisherQuillFormattingRange(position, rewrittenParagraph));
                originalPosition = checked(originalPosition + oldLength + 1);
                continue;
            }
            var uses = new List<PublisherQuillStyleUse>();
            int oldCursor = originalPosition;
            foreach (OfficeRichTextRun run in oldRuns) {
                PublisherQuillStyleUse? use = FindCharacter(oldCursor, Math.Max(1, run.Text.Length));
                if (use == null || run.Text.Length == 0 && oldRuns.Count != 1) return false;
                uses.Add(use.Value);
                oldCursor = checked(oldCursor + run.Text.Length);
            }
            PublisherQuillStyleUse? cr = FindCharacter(oldCursor, 1);
            if (cr == null) return false;
            // Removing all visible text must not silently substitute a different
            // native typing style for the requested empty paragraph's run.
            if (typingStyles == null && newRuns.Count == 1 && newRuns[0].Text.Length == 0
                && !Copy(cr.Value.DirectDefinition).SequenceEqual(Copy(uses[uses.Count - 1].DirectDefinition))) return false;
            for (int index = 0; index < newRuns.Count; index++) {
                token.ThrowIfCancellationRequested();
                if (newRuns[index].Text.Length == 0) continue;
                int original = oldRuns.Count == 1 ? 0 : index;
                byte[]? style = PublisherQuillCharacterStyleWriter.Rewrite(data, uses[original].DirectDefinition,
                    oldRuns[original], newRuns[index], context, resources);
                if (style == null) return false;
                position = checked(position + newRuns[index].Text.Length * 2);
                AppendRange(pc, position, style);
            }
            byte[]? terminator;
            if (newRuns.Count == 0) terminator = Copy(cr.Value.DirectDefinition);
            else {
                OfficeRichTextRun originalTyping = oldRuns[oldRuns.Count - 1];
                if (newRuns.Count == 1 && newRuns[0].Text.Length == 0 && typingStyles != null)
                    originalTyping = typingStyles.Run(string.Empty, typingStyles.CharacterAt(oldCursor)?.Style, typingStyles.ParagraphAt(oldCursor));
                terminator = PublisherQuillCharacterStyleWriter.Rewrite(data, cr.Value.DirectDefinition,
                    originalTyping, newRuns[newRuns.Count - 1], context, resources);
            }
            if (terminator == null) return false;
            position = checked(position + 2);
            AppendRange(pc, position, terminator);
            byte[]? definition = PublisherQuillParagraphStyleWriter.Rewrite(data, paragraph.DirectDefinition,
                before.Paragraphs[paragraphIndex], after.Paragraphs[paragraphIndex], context);
            if (definition == null) return false;
            pp.Add(new PublisherQuillFormattingRange(position, definition));
            originalPosition = checked(oldCursor + 1);
        }
        return originalPosition == binding.NativeStart + binding.NativeLength && position == outputTextOffset + after.Text.Length * 2;

        PublisherQuillStyleUse? FindCharacter(int start, int length) {
            foreach (PublisherQuillStyleUse use in characters)
                if (use.NativeStart <= start && (long)use.NativeStart + use.NativeLength >= (long)start + length) return use;
            return null;
        }
        byte[] Copy(PublisherQuillSourceSpan span) {
            var bytes = new byte[span.Length];
            Buffer.BlockCopy(data.Bytes, span.Offset, bytes, 0, bytes.Length);
            return bytes;
        }
    }

    private static void AppendRange(List<PublisherQuillFormattingRange> ranges, int end, byte[] style) {
        if (ranges.Count != 0 && ranges[ranges.Count - 1].Style.SequenceEqual(style))
            ranges[ranges.Count - 1] = new PublisherQuillFormattingRange(end, ranges[ranges.Count - 1].Style);
        else ranges.Add(new PublisherQuillFormattingRange(end, style));
    }

    private static byte[]? Pack(IReadOnlyList<PublisherQuillFormattingRange> ranges) =>
        ranges.Count is < 1 or > 84 || 8 + ranges.Count * 6 + ranges.Sum(range => range.Style.Length) > PublisherQuillFormattingPage.PageBytes
            ? null : PublisherQuillFormattingPage.Write(ranges);
}
