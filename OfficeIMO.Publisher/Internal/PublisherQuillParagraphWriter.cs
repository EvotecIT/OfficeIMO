using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Plans one character and paragraph page for an unchanged paragraph topology, retaining source OPLs and native CR formatting.</summary>
internal static class PublisherQuillParagraphWriter {
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
        out List<PublisherQuillFormattingRange> pc, out List<PublisherQuillFormattingRange> pp) {
        pc = new(); pp = new();
        // Local lists can be captured by the range-coalescing helper.
        var charactersOut = pc;
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
            if (newRuns.Count == 1 && newRuns[0].Text.Length == 0
                && !Copy(cr.Value.DirectDefinition).SequenceEqual(Copy(uses[uses.Count - 1].DirectDefinition))) return false;
            for (int index = 0; index < newRuns.Count; index++) {
                token.ThrowIfCancellationRequested();
                int original = oldRuns.Count == 1 ? 0 : index;
                byte[]? style = PublisherQuillCharacterStyleWriter.Rewrite(data, uses[original].DirectDefinition,
                    oldRuns[original], newRuns[index], context);
                if (style == null) return false;
                if (newRuns[index].Text.Length == 0) continue;
                position = checked(position + newRuns[index].Text.Length * 2);
                if (!Append(position, style)) return false;
            }
            byte[]? terminator = PublisherQuillCharacterStyleWriter.Rewrite(data, cr.Value.DirectDefinition,
                oldRuns[oldRuns.Count - 1], newRuns[newRuns.Count - 1], context);
            if (terminator == null) return false;
            position = checked(position + 2);
            if (!Append(position, terminator)) return false;
            pp.Add(new PublisherQuillFormattingRange(position, Copy(paragraph.DirectDefinition)));
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
        bool Append(int end, byte[] style) {
            // One original character range can span several paragraphs. Keep
            // identical adjacent definitions shared by one continuous range.
            if (charactersOut.Count != 0 && charactersOut[charactersOut.Count - 1].Style.SequenceEqual(style))
                charactersOut[charactersOut.Count - 1] = new PublisherQuillFormattingRange(end, charactersOut[charactersOut.Count - 1].Style);
            else charactersOut.Add(new PublisherQuillFormattingRange(end, style));
            return true;
        }
    }

    private static byte[]? Pack(IReadOnlyList<PublisherQuillFormattingRange> ranges) =>
        ranges.Count is < 1 or > 84 || 8 + ranges.Count * 6 + ranges.Sum(range => range.Style.Length) > PublisherQuillFormattingPage.PageBytes
            ? null : PublisherQuillFormattingPage.Write(ranges);
}
