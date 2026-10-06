using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    /// <summary>
    /// Separates table-owned words from adjacent prose after tagged-table enrichment.
    /// Physical line grouping predates that enrichment, so one line can span both owners.
    /// Keep both subsets traceable to the original runs and retain their local reading order.
    /// </summary>
    private static List<PdfUnderstandingLine> SplitTableProjectionLines(
        List<PdfUnderstandingLine> lines,
        IReadOnlyList<PdfUnderstandingTableCandidate> candidates,
        PdfUnderstandingPageResult analysis,
        PdfTextLayoutOptions? options,
        CancellationToken cancellationToken) {
        if (candidates.Count == 0) return lines;
        var tableRuns = new HashSet<PdfTextSpan>();
        var tableWords = new HashSet<PdfUnderstandingWord>();
        foreach (PdfUnderstandingTableCandidate candidate in candidates) {
            foreach (PdfUnderstandingLine line in candidate.SourceLines) {
                foreach (PdfUnderstandingWord word in line.Words) {
                    cancellationToken.ThrowIfCancellationRequested();
                    analysis.ConsumeWork?.Invoke(1);
                    tableWords.Add(word);
                    foreach (PdfTextSpan run in word.SourceRuns) {
                        analysis.ConsumeWork?.Invoke(1);
                        tableRuns.Add(run);
                    }
                }
            }
        }
        var result = new List<PdfUnderstandingLine>(lines.Count);
        foreach (PdfUnderstandingLine line in lines) {
            IReadOnlyList<PdfUnderstandingWord> projectionWords = SplitMixedSourceWords(line.Words);
            int segmentStart = 0;
            bool segmentIsTable = IsTableWord(projectionWords[0]);
            for (int wordIndex = 1; wordIndex < projectionWords.Count; wordIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                analysis.ConsumeWork?.Invoke(1);
                bool isTable = IsTableWord(projectionWords[wordIndex]);
                if (segmentIsTable == isTable) continue;
                AddSubset(wordIndex);
                segmentIsTable = isTable;
            }
            if (segmentStart == 0) result.Add(line);
            else AddSubset(projectionWords.Count);

            void AddSubset(int end) {
                PdfUnderstandingWord[] words = projectionWords.Skip(segmentStart).Take(end - segmentStart).ToArray();
                result.Add(PdfAdvancedUnderstandingStages.CreateLineSubset(line, words,
                    options?.ReadingDirection ?? PdfReadingDirection.Auto,
                    analysis.ConsumeWork, analysis.CancellationCheck));
                segmentStart = end;
            }
        }
        return result;

        // A custom grouping stage may merge table cells and prose into one word.
        // Partition that word by its original runs before applying line ownership.
        IReadOnlyList<PdfUnderstandingWord> SplitMixedSourceWords(IReadOnlyList<PdfUnderstandingWord> words) {
            List<PdfUnderstandingWord>? expanded = null;
            for (int wordIndex = 0; wordIndex < words.Count; wordIndex++) {
                PdfUnderstandingWord word = words[wordIndex];
                IReadOnlyList<PdfTextSpan> runs = word.SourceRuns;
                bool hasTable = false;
                bool hasProse = false;
                foreach (PdfTextSpan run in runs) {
                    cancellationToken.ThrowIfCancellationRequested();
                    analysis.ConsumeWork?.Invoke(1);
                    if (tableRuns.Contains(run)) hasTable = true;
                    else hasProse = true;
                }
                if (!hasTable || !hasProse) {
                    expanded?.Add(word);
                    continue;
                }
                expanded ??= words.Take(wordIndex).ToList();
                int start = 0;
                for (int end = 1; end <= runs.Count; end++) {
                    analysis.ConsumeWork?.Invoke(1);
                    if (end < runs.Count && tableRuns.Contains(runs[start]) == tableRuns.Contains(runs[end])) continue;
                    PdfTextSpan[] subset = runs.Skip(start).Take(end - start).ToArray();
                    List<TextLayoutEngine.TextLine> layout = TextLayoutEngine.BuildLines(subset,
                        new TextLayoutEngine.Options {
                            ForceSingleColumn = true,
                            ReadingDirection = options?.ReadingDirection ?? PdfReadingDirection.Auto
                        }, analysis.ConsumeWork, analysis.CancellationCheck);
                    expanded.Add(new PdfUnderstandingWord(
                        string.Join(" ", layout.Select(static item => item.Text)),
                        subset.Min(static run => run.X),
                        subset.Max(static run => run.X + Math.Max(0D, run.Advance)),
                        subset.Average(static run => run.Y),
                        subset.Max(static run => run.FontSize),
                        subset.Average(static run => run.RotationDegrees),
                        Array.AsReadOnly(subset), word.Confidence, word.Evidence,
                        subset.Sum(static run => Math.Max(0D, run.Advance)),
                        sourceSequence: word.SourceSequence) { IsSelectionBox = word.IsSelectionBox });
                    start = end;
                }
            }
            return expanded is null ? words : expanded;
        }

        bool IsTableWord(PdfUnderstandingWord word) {
            cancellationToken.ThrowIfCancellationRequested();
            if (tableWords.Contains(word)) return true;
            if (word.SourceRuns.Count == 0) return false;
            foreach (PdfTextSpan run in word.SourceRuns) {
                analysis.ConsumeWork?.Invoke(1);
                if (!tableRuns.Contains(run)) return false;
            }
            return true;
        }
    }
}
