using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Retains complete native paragraph definitions, resolving opaque ambiguity through immutable source identity.</summary>
    private static bool TryPlanRetention(PublisherBinaryData data, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int outputTextOffset, CancellationToken token,
        List<PublisherQuillFormattingRange> pc, List<PublisherQuillFormattingRange> pp) {
        var paragraphs = binding.Uses.Where(use => use.Tag == "FDPP").ToArray();
        var characters = binding.Uses.Where(use => use.Tag == "FDPC").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count || paragraphs.Length == 0 || after.Paragraphs.Count == 0) return false;
        var templates = new List<RetainedParagraph>(paragraphs.Length);
        int position = binding.NativeStart, characterIndex = 0;
        for (int i = 0; i < paragraphs.Length; i++) {
            token.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = before.Paragraphs[i];
            int length = paragraph.Runs.Sum(run => run.Text.Length), end = checked(position + length);
            if (length == 0 || paragraph.Label != null || paragraphs[i].NativeStart != position || paragraphs[i].NativeLength != length + 1) return false;
            var body = new List<PublisherQuillFormattingRange>();
            int cursor = position;
            while (cursor < end) {
                PublisherQuillStyleUse? use = At(cursor);
                if (use == null) return false;
                int next = (int)Math.Min(end, (long)use.Value.NativeStart + use.Value.NativeLength);
                AppendRange(body, checked((next - position) * 2), Copy(use.Value.DirectDefinition)); cursor = next;
            }
            PublisherQuillStyleUse? cr = At(end);
            if (cr == null) return false;
            templates.Add(new RetainedParagraph(paragraph, Copy(paragraphs[i].DirectDefinition), body, Copy(cr.Value.DirectDefinition), length));
            position = checked(end + 1);
        }
        if (position != binding.NativeStart + binding.NativeLength) return false;
        byte[] final = templates[templates.Count - 1].Terminator;
        RetainedParagraph lastSource = templates[templates.Count - 1];
        // A former final CR carries the story's typing style. Moving that paragraph
        // inside needs an observed body-equivalent interior CR, not the final CR.
        byte[]? movedTerminator = lastSource.Body.Count == 0 ? null : templates.Take(templates.Count - 1)
            .Select(item => item.Terminator).FirstOrDefault(style => style.SequenceEqual(lastSource.Body[lastSource.Body.Count - 1].Style));
        var references = new Dictionary<OfficeRichTextParagraph, RetainedParagraph>();
        var values = new Dictionary<OfficeRichTextParagraph, RetentionCandidates>(PublisherTextComparison.Paragraphs);
        for (int i = 0; i < templates.Count; i++) {
            RetainedParagraph template = templates[i];
            template.InteriorTerminator = i + 1 == templates.Count ? movedTerminator : template.Terminator;
            references.Add(template.Paragraph, template);
            if (!values.TryGetValue(template.Paragraph, out RetentionCandidates? candidates))
                values.Add(template.Paragraph, new RetentionCandidates(template));
            else candidates.Add(template);
        }
        position = outputTextOffset;
        for (int i = 0; i < after.Paragraphs.Count; i++) {
            token.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = after.Paragraphs[i];
            bool last = i + 1 == after.Paragraphs.Count;
            if (!references.TryGetValue(paragraph, out RetainedParagraph? template)) {
                if (!values.TryGetValue(paragraph, out RetentionCandidates? candidates) || !candidates.IsUnambiguous(last)) return false;
                template = candidates.First;
            }
            byte[]? terminator = last ? final : template.InteriorTerminator;
            if (terminator == null) return false;
            foreach (var range in template.Body) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
            position = checked(position + (template.Length + 1) * 2);
            AppendRange(pc, position, terminator);
            pp.Add(new PublisherQuillFormattingRange(position, template.Definition));
        }
        return position == outputTextOffset + after.Text.Length * 2;

        PublisherQuillStyleUse? At(int value) {
            while (characterIndex < characters.Length && (long)characters[characterIndex].NativeStart + characters[characterIndex].NativeLength <= value)
                characterIndex++;
            return characterIndex < characters.Length && characters[characterIndex].NativeStart <= value ? characters[characterIndex] : null;
        }
        byte[] Copy(PublisherQuillSourceSpan span) {
            var bytes = new byte[span.Length]; Buffer.BlockCopy(data.Bytes, span.Offset, bytes, 0, bytes.Length); return bytes;
        }
    }

    private sealed class RetainedParagraph {
        internal RetainedParagraph(OfficeRichTextParagraph paragraph, byte[] definition, List<PublisherQuillFormattingRange> body, byte[] terminator, int length) {
            Paragraph = paragraph; Definition = definition; Body = body; Terminator = terminator; Length = length;
        }
        internal OfficeRichTextParagraph Paragraph { get; }
        internal byte[] Definition { get; }
        internal List<PublisherQuillFormattingRange> Body { get; }
        internal byte[] Terminator { get; }
        internal byte[]? InteriorTerminator { get; set; }
        internal int Length { get; }
    }

    private sealed class RetentionCandidates {
        internal RetentionCandidates(RetainedParagraph first) { First = first; }
        internal RetainedParagraph First { get; }
        private bool _sameBody = true, _sameInterior = true;
        internal void Add(RetainedParagraph other) {
            _sameBody &= First.Definition.SequenceEqual(other.Definition) && First.Body.Count == other.Body.Count
                && First.Body.Zip(other.Body, (a, b) => a.AbsoluteEnd == b.AbsoluteEnd && a.Style.SequenceEqual(b.Style)).All(equal => equal);
            _sameInterior &= First.InteriorTerminator != null && other.InteriorTerminator != null
                && First.InteriorTerminator.SequenceEqual(other.InteriorTerminator);
        }
        internal bool IsUnambiguous(bool final) => _sameBody && (final || _sameInterior && First.InteriorTerminator != null);
    }
}
