namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Count changes repeat unambiguous source definitions; projected equality alone does not establish native equivalence.</summary>
    private static bool TryPlanTopology(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int outputTextOffset, CancellationToken token,
        List<PublisherQuillFormattingRange> pc, List<PublisherQuillFormattingRange> pp, PublisherQuillWriteResources? resources) {
        var paragraphs = binding.Uses.Where(use => use.Tag == "FDPP").ToArray();
        var characters = binding.Uses.Where(use => use.Tag == "FDPC").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count || paragraphs.Length == 0 || after.Paragraphs.Count == 0) return false;
        PublisherQuillStyleUse? body = At(binding.NativeStart), final = At(binding.NativeStart + binding.NativeLength - 1);
        if (body == null || final == null) return false;
        byte[] bodyBytes = Copy(body.Value.DirectDefinition), paragraphBytes = Copy(paragraphs[0].DirectDefinition);
        int originalPosition = binding.NativeStart;
        for (int i = 0; i < before.Paragraphs.Count; i++) {
            token.ThrowIfCancellationRequested();
            int length = before.Paragraphs[i].Runs.Sum(run => run.Text.Length);
            if (before.Paragraphs[i].Runs.Count != 1 || paragraphs[i].NativeStart != originalPosition
                || paragraphs[i].NativeLength != length + 1 || !Copy(paragraphs[i].DirectDefinition).SequenceEqual(paragraphBytes)) return false;
            int visibleEnd = checked(originalPosition + length);
            foreach (PublisherQuillStyleUse use in characters.Where(use => use.NativeStart < visibleEnd
                && (long)use.NativeStart + use.NativeLength > originalPosition))
                if (!Copy(use.DirectDefinition).SequenceEqual(bodyBytes)) return false;
            if (i + 1 < before.Paragraphs.Count) {
                PublisherQuillStyleUse? cr = At(visibleEnd);
                if (cr == null || !Copy(cr.Value.DirectDefinition).SequenceEqual(bodyBytes)) return false;
            }
            originalPosition = checked(visibleEnd + 1);
        }
        if (originalPosition != binding.NativeStart + binding.NativeLength) return false;
        var originalRun = before.Paragraphs[0].Runs[0]; int position = outputTextOffset;
        for (int i = 0; i < after.Paragraphs.Count; i++) {
            token.ThrowIfCancellationRequested(); var runs = after.Paragraphs[i].Runs;
            bool last = i + 1 == after.Paragraphs.Count;
            if (runs.Count == 0 || last && runs.All(run => run.Text.Length == 0)
                && !Copy(final.Value.DirectDefinition).SequenceEqual(bodyBytes)) return false;
            foreach (var run in runs) {
                byte[]? style = PublisherQuillCharacterStyleWriter.Rewrite(data, body.Value.DirectDefinition, originalRun, run, context, resources);
                if (style == null) return false;
                if (run.Text.Length == 0) continue;
                position = checked(position + run.Text.Length * 2); AppendRange(pc, position, style);
            }
            byte[]? terminator = PublisherQuillCharacterStyleWriter.Rewrite(data,
                last ? final.Value.DirectDefinition : body.Value.DirectDefinition, originalRun, runs[runs.Count - 1], context, resources);
            if (terminator == null) return false;
            position = checked(position + 2); AppendRange(pc, position, terminator);
            pp.Add(new PublisherQuillFormattingRange(position, paragraphBytes));
        }
        return position == outputTextOffset + after.Text.Length * 2;

        PublisherQuillStyleUse? At(int value) {
            foreach (PublisherQuillStyleUse use in characters)
                if (use.NativeStart <= value && (long)use.NativeStart + use.NativeLength > value) return use;
            return null;
        }
        byte[] Copy(PublisherQuillSourceSpan span) {
            var bytes = new byte[span.Length]; Buffer.BlockCopy(data.Bytes, span.Offset, bytes, 0, bytes.Length); return bytes;
        }
    }
}
