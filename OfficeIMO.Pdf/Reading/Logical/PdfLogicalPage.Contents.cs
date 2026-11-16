namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    /// <summary>
    /// Projects recognised contents entries as list items that keep their link and printed page label, whether
    /// the entry came from TOC/TOCI structure or from the untagged contents-line pattern.
    /// </summary>
    private static void AddContentsItems(
        int pageNumber,
        List<PdfLogicalTextBlock> textBlocks,
        PdfUnderstandingPageResult analysis,
        IReadOnlyList<PdfLogicalLinkAnnotation> links,
        List<PdfLogicalListItem> result,
        HashSet<PdfLogicalTextBlock> represented) {
        if (analysis.ContentsEntries.Count == 0) return;
        var owners = new Dictionary<PdfTextSpan, PdfContentsEntrySource>();
        foreach (PdfContentsEntrySource entry in analysis.ContentsEntries) {
            foreach (PdfTextSpan run in entry.Runs) {
                Consume();
                owners[run] = entry;
            }
        }

        var blocksByEntry = new Dictionary<PdfContentsEntrySource, List<PdfLogicalTextBlock>>();
        foreach (PdfLogicalTextBlock block in textBlocks) {
            Consume();
            if (block.Spans.Count == 0 || !owners.TryGetValue(block.Spans[0], out PdfContentsEntrySource? entry) ||
                !block.Spans.All(span => owners.TryGetValue(span, out PdfContentsEntrySource? owner) && ReferenceEquals(owner, entry))) continue;
            if (!blocksByEntry.TryGetValue(entry, out List<PdfLogicalTextBlock>? blocks)) {
                blocks = new List<PdfLogicalTextBlock>();
                blocksByEntry.Add(entry, blocks);
            }
            blocks.Add(block);
        }

        foreach (PdfContentsEntrySource entry in analysis.ContentsEntries) {
            Consume();
            if (!blocksByEntry.TryGetValue(entry, out List<PdfLogicalTextBlock>? blocks)) continue;
            // Caller-supplied stages can retain only part of an entry; keep the title within the projected lines.
            var retained = new HashSet<PdfTextSpan>(blocks.SelectMany(static block => block.Spans));
            PdfUnderstandingWord[] titleWords = entry.Title.Where(word => word.SourceRuns.All(retained.Contains)).ToArray();
            var runs = new List<PdfLogicalTextRun>();
            var builder = new System.Text.StringBuilder();
            PdfUnderstandingWord? previous = null;
            for (int index = 0; index < titleWords.Length; index++) {
                PdfUnderstandingWord word = titleWords[index];
                bool trimmed = entry.LastTitleText is not null && ReferenceEquals(word, entry.Title[entry.Title.Count - 1]);
                string text = trimmed ? entry.LastTitleText! : word.Text.Trim();
                if (text.Length == 0) continue;
                // Ligature and style-change pieces of one word follow each other without a space.
                if (previous is not null && !PdfDocumentSemanticEnricher.ContinuesContentsWord(previous, word)) {
                    runs.Add(new PdfLogicalTextRun(" ", sourceSpan: null));
                    builder.Append(' ');
                }
                AddWordRuns(runs, word, text);
                builder.Append(text);
                previous = word;
            }
            if (builder.Length == 0) continue;
            string title = builder.ToString();
            string[] lineTexts = blocks.Select((_, index) => index == 0 ? title : string.Empty).ToArray();
            result.Add(new PdfLogicalListItem(pageNumber, entry.Level, string.Empty, title, blocks, lineTexts, 0.97D,
                new[] { entry.Evidence }, runs) {
                IsContentsEntry = true,
                TargetPageLabel = entry.PageLabel,
                Link = FindTitleLink(entry, links)
            });
            foreach (PdfLogicalTextBlock block in blocks) represented.Add(block);
        }

        void Consume() {
            analysis.CancellationCheck?.Invoke();
            analysis.ConsumeWork?.Invoke(1);
        }
    }

    /// <summary>
    /// Keeps each source span's style within a word, such as bold "Office" followed by regular "IMO". When the
    /// spans do not spell the word, the word keeps its first span's style.
    /// </summary>
    private static void AddWordRuns(List<PdfLogicalTextRun> runs, PdfUnderstandingWord word, string text) {
        if (word.SourceRuns.Count > 1) {
            string joined = string.Concat(word.SourceRuns.Select(static span => span.Text.Trim()));
            if (joined.StartsWith(text, StringComparison.Ordinal)) {
                int remaining = text.Length;
                foreach (PdfTextSpan span in word.SourceRuns) {
                    if (remaining == 0) break;
                    string piece = span.Text.Trim();
                    if (piece.Length == 0) continue;
                    if (piece.Length > remaining) piece = piece.Substring(0, remaining);
                    runs.Add(new PdfLogicalTextRun(piece, span));
                    remaining -= piece.Length;
                }
                return;
            }
        }
        runs.Add(new PdfLogicalTextRun(text, word.SourceRuns.Count > 0 ? word.SourceRuns[0] : null));
    }

    private static PdfLogicalLinkAnnotation? FindTitleLink(PdfContentsEntrySource entry, IReadOnlyList<PdfLogicalLinkAnnotation> links) {
        PdfLogicalLinkAnnotation? best = null;
        double bestOverlap = 0D;
        foreach (PdfLogicalLinkAnnotation link in links) {
            double left = Math.Min(link.X1, link.X2);
            double right = Math.Max(link.X1, link.X2);
            double bottom = Math.Min(link.Y1, link.Y2);
            double top = Math.Max(link.Y1, link.Y2);
            double overlap = 0D;
            foreach (PdfUnderstandingWord word in entry.Title) {
                double wordTop = word.BaselineY + word.FontSize * 0.7D;
                double wordBottom = word.BaselineY - word.FontSize * 0.2D;
                if (wordBottom > top || wordTop < bottom) continue;
                overlap += Math.Max(0D, Math.Min(right, word.XEnd) - Math.Max(left, word.XStart));
            }
            if (overlap > bestOverlap) {
                bestOverlap = overlap;
                best = link;
            }
        }
        return best;
    }
}
