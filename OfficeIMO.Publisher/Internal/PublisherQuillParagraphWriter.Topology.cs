using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Count changes repeat unambiguous source definitions; projected equality alone does not establish native equivalence.</summary>
    private static bool TryPlanTopology(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int outputTextOffset, CancellationToken token,
        List<PublisherQuillFormattingRange> pc, List<PublisherQuillFormattingRange> pp, PublisherQuillWriteResources? resources,
        PublisherQuillStyles? typingStyles) {
        var paragraphs = binding.Uses.Where(use => use.Tag == "FDPP").ToArray();
        var characters = binding.Uses.Where(use => use.Tag == "FDPC").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count || paragraphs.Length == 0 || after.Paragraphs.Count == 0) return false;
        PublisherQuillStyleUse? body = At(binding.NativeStart), final = At(binding.NativeStart + binding.NativeLength - 1);
        if (body == null || final == null) return false;
        byte[] bodyBytes = Copy(body.Value.DirectDefinition), paragraphBytes = Copy(paragraphs[0].DirectDefinition);
        var emptyReferences = new Dictionary<OfficeRichTextParagraph, (byte[] Bytes, OfficeRichTextRun Run)>();
        var emptyValues = new Dictionary<OfficeRichTextParagraph, (byte[] Bytes, OfficeRichTextRun Run)?>(PublisherTextComparison.Paragraphs);
        int originalPosition = binding.NativeStart;
        for (int i = 0; i < before.Paragraphs.Count; i++) {
            token.ThrowIfCancellationRequested();
            int length = before.Paragraphs[i].Runs.Sum(run => run.Text.Length);
            if (before.Paragraphs[i].Runs.Count != 1 || paragraphs[i].NativeStart != originalPosition
                || paragraphs[i].NativeLength != length + 1 || !Copy(paragraphs[i].DirectDefinition).SequenceEqual(paragraphBytes)) return false;
            int visibleEnd = checked(originalPosition + length);
            if (length == 0) {
                PublisherQuillStyleUse? source = At(visibleEnd);
                if (source == null) return false;
                byte[] bytes = Copy(source.Value.DirectDefinition);
                OfficeRichTextParagraph paragraph = before.Paragraphs[i];
                var retained = (Bytes: bytes, Run: paragraph.Runs[0]);
                emptyReferences.Add(paragraph, retained);
                if (!emptyValues.TryGetValue(paragraph, out var prior)) emptyValues.Add(paragraph, retained);
                else if (prior.HasValue && !prior.Value.Bytes.SequenceEqual(bytes)) emptyValues[paragraph] = null;
            }
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
            (byte[] Bytes, OfficeRichTextRun Run)? retainedEmpty = null;
            if (runs.Count == 1 && runs[0].Text.Length == 0) {
                if (emptyReferences.TryGetValue(after.Paragraphs[i], out var exact)) retainedEmpty = exact;
                else if (emptyValues.TryGetValue(after.Paragraphs[i], out retainedEmpty) && retainedEmpty == null) return false;
            }
            if (retainedEmpty == null && typingStyles == null && last && runs.Count == 1 && runs[0].Text.Length == 0
                && !Copy(final.Value.DirectDefinition).SequenceEqual(bodyBytes)) return false;
            foreach (var run in runs) {
                if (run.Text.Length == 0) continue;
                byte[]? style = PublisherQuillCharacterStyleWriter.Rewrite(data, body.Value.DirectDefinition, originalRun, run, context, resources);
                if (style == null) return false;
                position = checked(position + run.Text.Length * 2); AppendRange(pc, position, style);
            }
            PublisherQuillSourceSpan ending = last ? final.Value.DirectDefinition : body.Value.DirectDefinition;
            byte[]? terminator;
            // A retained empty CR must survive even when other paragraphs are
            // newly authored; uniform formatting does not establish its identity.
            if (retainedEmpty is { } retained) {
                terminator = PublisherQuillCharacterStyleWriter.RewriteExplicitAttributes(retained.Bytes, retained.Run, runs[0], context);
            }
            else if (runs.Count == 0) terminator = Copy(ending);
            else {
                var originalTyping = originalRun;
                if (last && runs.Count == 1 && runs[0].Text.Length == 0 && typingStyles != null) {
                    int finalPosition = binding.NativeStart + binding.NativeLength - 1;
                    originalTyping = typingStyles.Run(string.Empty, typingStyles.CharacterAt(finalPosition)?.Style, typingStyles.ParagraphAt(finalPosition));
                }
                terminator = PublisherQuillCharacterStyleWriter.Rewrite(data, ending, originalTyping, runs[runs.Count - 1], context, resources);
            }
            if (terminator == null) return false;
            position = checked(position + 2); AppendRange(pc, position, terminator);
            byte[]? definition = PublisherQuillParagraphStyleWriter.Rewrite(data, paragraphs[0].DirectDefinition,
                before.Paragraphs[0], after.Paragraphs[i], context);
            if (definition == null) return false;
            pp.Add(new PublisherQuillFormattingRange(position, definition));
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
