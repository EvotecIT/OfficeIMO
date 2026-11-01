using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    /// <summary>Retains native templates, resolving exact identity before retained-text formatting edits and format-matched text replacement.</summary>
    private static bool TryPlanRetention(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, int outputTextOffset, CancellationToken token,
        List<PublisherQuillFormattingRange> pc, List<PublisherQuillFormattingRange> pp, PublisherQuillWriteResources? resources) {
        var paragraphs = binding.Uses.Where(use => use.Tag == "FDPP").ToArray();
        var characters = binding.Uses.Where(use => use.Tag == "FDPC").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count || paragraphs.Length == 0 || after.Paragraphs.Count == 0) return false;
        var templates = new List<RetainedParagraph>(paragraphs.Length);
        int position = binding.NativeStart, characterIndex = 0;
        for (int i = 0; i < paragraphs.Length; i++) {
            token.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = before.Paragraphs[i];
            int length = paragraph.Runs.Sum(run => run.Text.Length), end = checked(position + length);
            if (paragraph.Label != null || paragraphs[i].NativeStart != position || paragraphs[i].NativeLength != length + 1) return false;
            var body = new List<PublisherQuillFormattingRange>();
            var bodySpans = new List<PublisherQuillSourceSpan>();
            int cursor = position;
            while (cursor < end) {
                PublisherQuillStyleUse? use = At(cursor);
                if (use == null) return false;
                int next = (int)Math.Min(end, (long)use.Value.NativeStart + use.Value.NativeLength);
                int count = body.Count;
                AppendRange(body, checked((next - position) * 2), Copy(use.Value.DirectDefinition));
                if (body.Count != count) bodySpans.Add(use.Value.DirectDefinition);
                cursor = next;
            }
            PublisherQuillStyleUse? cr = At(end);
            if (cr == null) return false;
            templates.Add(new RetainedParagraph(paragraph, paragraphs[i].DirectDefinition, Copy(paragraphs[i].DirectDefinition),
                body, bodySpans, cr.Value.DirectDefinition, Copy(cr.Value.DirectDefinition), length, position));
            position = checked(end + 1);
        }
        if (position != binding.NativeStart + binding.NativeLength) return false;
        byte[] final = templates[templates.Count - 1].Terminator;
        RetainedParagraph lastSource = templates[templates.Count - 1];
        // A former final CR carries the story's typing style. Moving that paragraph
        // inside needs an observed body-equivalent interior CR, not the final CR.
        RetainedParagraph? moved = lastSource.Length == 0 ? lastSource : templates.Take(templates.Count - 1)
            .FirstOrDefault(item => item.Terminator.SequenceEqual(lastSource.Body[lastSource.Body.Count - 1].Style));
        byte[]? movedTerminator = moved?.Terminator;
        var references = new Dictionary<OfficeRichTextParagraph, RetainedParagraph>();
        var values = new Dictionary<OfficeRichTextParagraph, RetentionCandidates>(PublisherTextComparison.Paragraphs);
        var replacements = new Dictionary<OfficeRichTextParagraph, TextReplacementCandidates>(PublisherTextComparison.ParagraphFormats);
        var formatting = new Dictionary<OfficeRichTextParagraph, TextReplacementCandidates>(PublisherTextComparison.ParagraphTexts);
        var content = new Dictionary<OfficeRichTextParagraph, RetentionCandidates>(PublisherTextComparison.ParagraphContent);
        for (int i = 0; i < templates.Count; i++) {
            RetainedParagraph template = templates[i];
            template.InteriorTerminator = i + 1 == templates.Count ? movedTerminator : template.Terminator;
            references.Add(template.Paragraph, template);
            if (!values.TryGetValue(template.Paragraph, out RetentionCandidates? candidates))
                values.Add(template.Paragraph, new RetentionCandidates(template));
            else candidates.Add(template);
            if (template.Length > 0 && template.Paragraph.Runs.All(r => r.Text.Length > 0)) {
                if (!content.TryGetValue(template.Paragraph, out RetentionCandidates? matchingContent))
                    content.Add(template.Paragraph, new RetentionCandidates(template));
                else matchingContent.Add(template);
            }
            if (template.Paragraph.Runs.Count > 0) {
                if (!replacements.TryGetValue(template.Paragraph, out TextReplacementCandidates? matching))
                    replacements.Add(template.Paragraph, new TextReplacementCandidates(template));
                else matching.Add(template);
                if (!formatting.TryGetValue(template.Paragraph, out TextReplacementCandidates? retainedText))
                    formatting.Add(template.Paragraph, new TextReplacementCandidates(template));
                else retainedText.Add(template);
            }
        }
        position = outputTextOffset;
        for (int i = 0; i < after.Paragraphs.Count; i++) {
            token.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = after.Paragraphs[i];
            bool last = i + 1 == after.Paragraphs.Count;
            bool replaceText = false, reformat = false, resegment = false;
            if (!references.TryGetValue(paragraph, out RetainedParagraph? template)) {
                if (values.TryGetValue(paragraph, out RetentionCandidates? candidates)) {
                    if (!candidates.IsUnambiguous(last)) return false;
                    template = candidates.First;
                } else {
                    if (paragraph.Runs.Count == 0 || paragraph.Runs.Any(r => r.Text.Length == 0)) return false;
                    // Retained text is stronger correspondence than a requested
                    // style that may belong to a different native paragraph.
                    if (formatting.TryGetValue(paragraph, out TextReplacementCandidates? retainedText)) {
                        if (!retainedText.IsUnambiguous(last)) return false;
                        template = retainedText.First; reformat = true;
                    } else if (content.TryGetValue(paragraph, out RetentionCandidates? retainedContent)) {
                        if (!retainedContent.IsUnambiguous(last)) return false;
                        template = retainedContent.First; resegment = true;
                    } else {
                        if (!replacements.TryGetValue(paragraph, out TextReplacementCandidates? matching)) {
                            if (!TryPlanEmptyRefill(data, context, templates, paragraph, token, resources, out var refill, out var refillDefinition)) return false;
                            foreach (var range in refill) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
                            position = checked(position + (paragraph.Runs.Sum(r => r.Text.Length) + 1) * 2);
                            pp.Add(new PublisherQuillFormattingRange(position, refillDefinition!));
                            continue;
                        }
                        if (!matching.IsUnambiguous(last)) return false;
                        template = matching.First; replaceText = true;
                    }
                }
            }
            // Visual correspondence selects the native template; explicit
            // nullable attribute requests still need the character writer on copied bodies.
            bool attributesChanged = !resegment && template.Paragraph.Runs.Count == paragraph.Runs.Count
                && template.Paragraph.Runs.Zip(paragraph.Runs, PublisherQuillCharacterStyleWriter.HasExplicitAttributeChange).Any(value => value);
            if (attributesChanged && (template.Length != 0 || replaceText)) {
                if (template.ReplacementSpans != null) reformat = true;
                else if (!replaceText) resegment = true;
                else return false;
            }
            // The CR is an empty paragraph's only character definition. It must
            // travel with that paragraph, including its typing and opaque fields.
            byte[]? terminator = template.Length == 0 ? template.Terminator : last ? final : template.InteriorTerminator;
            if (terminator == null) return false;
            if (paragraph.Runs.Count > 0 && template.Paragraph.Runs.Count > 0) {
                terminator = PublisherQuillCharacterStyleWriter.RewriteExplicitAttributes(terminator,
                    template.Paragraph.Runs[template.Paragraph.Runs.Count - 1], paragraph.Runs[paragraph.Runs.Count - 1], context);
                if (terminator == null) return false;
            }
            byte[] definition = template.Definition;
            if (reformat || resegment) {
                byte[]? rewritten = PublisherQuillParagraphStyleWriter.Rewrite(data, template.DefinitionSpan, template.Paragraph, paragraph, context);
                if (rewritten == null) return false;
                definition = rewritten;
            }
            if (resegment) {
                if (!TryPlanContentRanges(data, context, characters, template.NativeStart, template.Paragraph, paragraph, token, resources, out var mapped)) return false;
                foreach (var range in mapped) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
                position = checked(position + (template.Length + 1) * 2);
            } else if (replaceText || reformat) {
                for (int run = 0; run < paragraph.Runs.Count; run++) {
                    byte[]? style = reformat ? PublisherQuillCharacterStyleWriter.Rewrite(data, template.ReplacementSpans![run],
                        template.Paragraph.Runs[run], paragraph.Runs[run], context, resources) : template.ReplacementBody![run];
                    if (style == null) return false;
                    position = checked(position + paragraph.Runs[run].Text.Length * 2);
                    AppendRange(pc, position, style);
                }
                position = checked(position + 2);
            } else {
                foreach (var range in template.Body) AppendRange(pc, checked(position + range.AbsoluteEnd), range.Style);
                position = checked(position + (template.Length + 1) * 2);
            }
            AppendRange(pc, position, terminator);
            pp.Add(new PublisherQuillFormattingRange(position, definition));
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
        internal RetainedParagraph(OfficeRichTextParagraph paragraph, PublisherQuillSourceSpan definitionSpan, byte[] definition,
            List<PublisherQuillFormattingRange> body, IReadOnlyList<PublisherQuillSourceSpan> bodySpans,
            PublisherQuillSourceSpan terminatorSpan, byte[] terminator, int length, int nativeStart) {
            Paragraph = paragraph; Definition = definition; Body = body; Terminator = terminator; Length = length;
            DefinitionSpan = definitionSpan; TerminatorSpan = terminatorSpan;
            NativeStart = nativeStart;
            ReplacementBody = MapRuns(bodySpans, terminatorSpan, out var mappedSpans); ReplacementSpans = mappedSpans;
        }
        internal OfficeRichTextParagraph Paragraph { get; }
        internal byte[] Definition { get; }
        internal PublisherQuillSourceSpan DefinitionSpan { get; }
        internal PublisherQuillSourceSpan TerminatorSpan { get; }
        internal List<PublisherQuillFormattingRange> Body { get; }
        internal byte[] Terminator { get; }
        internal byte[]? InteriorTerminator { get; set; }
        internal int Length { get; }
        internal int NativeStart { get; }
        internal IReadOnlyList<byte[]>? ReplacementBody { get; }
        internal IReadOnlyList<PublisherQuillSourceSpan>? ReplacementSpans { get; }

        // A public run may start inside a native range, but must not cross a
        // native definition boundary when its text length changes.
        private IReadOnlyList<byte[]>? MapRuns(IReadOnlyList<PublisherQuillSourceSpan> bodySpans, PublisherQuillSourceSpan terminatorSpan,
            out IReadOnlyList<PublisherQuillSourceSpan>? mappedSpans) {
            // Empty paragraphs have one typing run represented solely by their
            // native CR. Filling that run retains its complete definition.
            if (Length == 0 && Paragraph.Runs.Count == 1 && Paragraph.Runs[0].Text.Length == 0) {
                mappedSpans = new[] { terminatorSpan };
                return new[] { Terminator };
            }
            var result = new List<byte[]>(Paragraph.Runs.Count);
            var spans = new List<PublisherQuillSourceSpan>(Paragraph.Runs.Count); mappedSpans = null;
            int position = 0, native = 0;
            foreach (var run in Paragraph.Runs) {
                if (run.Text.Length == 0) return null;
                while (native < Body.Count && Body[native].AbsoluteEnd <= position) native++;
                int end = checked(position + run.Text.Length * 2);
                if (native >= Body.Count || end > Body[native].AbsoluteEnd) return null;
                result.Add(Body[native].Style); spans.Add(bodySpans[native]); position = end;
            }
            if (result.Count == 0 || position != checked(Length * 2)) return null;
            mappedSpans = spans; return result;
        }
    }

    private sealed class RetentionCandidates {
        internal RetentionCandidates(RetainedParagraph first) { First = first; }
        internal RetainedParagraph First { get; }
        private bool _sameBody = true, _sameInterior = true;
        internal void Add(RetainedParagraph other) {
            _sameBody &= First.Definition.SequenceEqual(other.Definition) && First.Body.Count == other.Body.Count
                && (First.Length != 0 || First.Terminator.SequenceEqual(other.Terminator))
                && First.Body.Zip(other.Body, (a, b) => a.AbsoluteEnd == b.AbsoluteEnd && a.Style.SequenceEqual(b.Style)).All(equal => equal);
            _sameInterior &= First.InteriorTerminator != null && other.InteriorTerminator != null
                && First.InteriorTerminator.SequenceEqual(other.InteriorTerminator);
        }
        internal bool IsUnambiguous(bool final) => _sameBody && (final || _sameInterior && First.InteriorTerminator != null);
    }

    // Replacement runs have new endpoints. Each corresponding source run needs
    // a uniform native definition, and all same-format candidates must agree.
    private sealed class TextReplacementCandidates {
        internal TextReplacementCandidates(RetainedParagraph first) {
            First = first; _sameBody = first.ReplacementBody != null;
        }
        internal RetainedParagraph First { get; }
        private bool _sameBody, _sameInterior = true;
        internal void Add(RetainedParagraph other) {
            _sameBody &= First.ReplacementBody != null && other.ReplacementBody != null
                && First.Definition.SequenceEqual(other.Definition) && First.ReplacementBody.Count == other.ReplacementBody.Count
                && First.ReplacementBody.Zip(other.ReplacementBody, (a, b) => a.SequenceEqual(b)).All(equal => equal);
            _sameInterior &= First.InteriorTerminator != null && other.InteriorTerminator != null
                && First.InteriorTerminator.SequenceEqual(other.InteriorTerminator);
        }
        internal bool IsUnambiguous(bool final) => _sameBody && (final || _sameInterior && First.InteriorTerminator != null);
    }
}
