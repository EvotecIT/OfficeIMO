namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    private static void AddTaggedListItems(
        int pageNumber,
        List<PdfLogicalTextBlock> textBlocks,
        Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement> semantics,
        PdfUnderstandingPageResult analysis,
        List<PdfLogicalListItem> result,
        HashSet<PdfLogicalTextBlock> represented) {
        if (analysis.TaggedListItems.Count == 0) return;
        var owners = new Dictionary<PdfTextSpan, (PdfTaggedListItemSource Item, bool Label)>();
        var groups = new Dictionary<PdfTaggedListItemSource, TaggedListProjection>();
        foreach (PdfTaggedListItemSource item in analysis.TaggedListItems) {
            foreach (PdfTextSpan run in item.BodyRuns) { Consume(); owners.Add(run, (item, false)); }
            foreach (PdfTextSpan run in item.LabelRuns) { Consume(); owners.Add(run, (item, true)); }
        }
        foreach (PdfLogicalTextBlock block in textBlocks) {
            Consume();
            var fragments = new Dictionary<PdfTaggedListItemSource, TaggedListLine>();
            int cursor = 0;
            bool exclusive = !block.IsTableContent && block.Kind is not (PdfLogicalElementKind.Heading or PdfLogicalElementKind.Table);
            foreach (PdfTextSpan span in block.Spans) {
                Consume();
                if (!PdfLogicalTextBlock.TryFindNormalizedSpan(block.Text, span.Text, cursor, out int start, out int end)) continue;
                cursor = end;
                if (!owners.TryGetValue(span, out var owner)) { exclusive = false; continue; }
                if (!fragments.TryGetValue(owner.Item, out TaggedListLine? fragment)) {
                    fragment = new TaggedListLine(block);
                    fragments.Add(owner.Item, fragment);
                }
                (owner.Label ? fragment.Label : fragment.Body).Append(block.Text, start, end);
            }
            foreach (var pair in fragments) {
                Consume();
                if (!groups.TryGetValue(pair.Key, out TaggedListProjection? group)) {
                    group = new TaggedListProjection();
                    groups.Add(pair.Key, group);
                }
                group.Lines.Add(pair.Value);
                group.CanProject &= exclusive && fragments.Count == 1;
            }
        }
        foreach (var pair in groups) {
            Consume();
            TaggedListProjection group = pair.Value;
            string[] body = group.Lines.Select(static line => line.Body.Text).ToArray();
            if (body.All(static text => text.Length == 0)) continue;
            string marker = string.Join(" ", group.Lines.Select(static line => line.Label.Text).Where(static text => text.Length > 0));
            if (marker.Length == 0) {
                int firstBody = Array.FindIndex(body, static text => text.Length > 0);
                if (ContentStructureExtractor.TryParseListItemText(body[firstBody], out string inferredMarker, out string text, out _)) {
                    marker = inferredMarker;
                    body[firstBody] = text;
                }
            }
            PdfLogicalTextBlock[] lines = group.Lines.Select(static line => line.Block).ToArray();
            result.Add(new PdfLogicalListItem(pageNumber, pair.Key.Level, marker,
                string.Join(" ", body.Where(static text => text.Length > 0)), lines, body, 0.98D,
                lines.Where(semantics.ContainsKey).SelectMany(line => semantics[line].Evidence).Distinct().Concat(
                new[] { new PdfInferenceEvidence("semantic.tagged-list-owner",
                    "The nearest tagged LI structure " + pair.Key.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " associates its label and body source runs on this page.", 0.99D) })) {
                CanProjectAsList = group.CanProject
            });
            foreach (PdfLogicalTextBlock line in lines) represented.Add(line);
        }

        void Consume() {
            analysis.CancellationCheck?.Invoke();
            analysis.ConsumeWork?.Invoke(1);
        }
    }

    private sealed class TaggedListProjection {
        internal List<TaggedListLine> Lines { get; } = new();
        internal bool CanProject { get; set; } = true;
    }

    private sealed class TaggedListLine {
        internal TaggedListLine(PdfLogicalTextBlock block) { Block = block; }
        internal PdfLogicalTextBlock Block { get; }
        internal TaggedListText Body { get; } = new();
        internal TaggedListText Label { get; } = new();
    }

    private sealed class TaggedListText {
        private readonly System.Text.StringBuilder _text = new();
        private int _lastEnd = -1;
        internal string Text => _text.ToString().Trim();

        internal void Append(string line, int start, int end) {
            if (_lastEnd >= 0 && start > _lastEnd) {
                string gap = line.Substring(_lastEnd, start - _lastEnd);
                _text.Append(string.IsNullOrWhiteSpace(gap) ? gap : " ");
            }
            _text.Append(line, start, end - start);
            _lastEnd = end;
        }
    }
}
