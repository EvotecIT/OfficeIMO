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
            int segmentStart = 0;
            bool segmentIsTable = IsTableWord(line.Words[0]);
            for (int wordIndex = 1; wordIndex < line.Words.Count; wordIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                analysis.ConsumeWork?.Invoke(1);
                bool isTable = IsTableWord(line.Words[wordIndex]);
                if (segmentIsTable == isTable) continue;
                AddSubset(wordIndex);
                segmentIsTable = isTable;
            }
            if (segmentStart == 0) result.Add(line);
            else AddSubset(line.Words.Count);

            void AddSubset(int end) {
                PdfUnderstandingWord[] words = line.Words.Skip(segmentStart).Take(end - segmentStart).ToArray();
                result.Add(PdfAdvancedUnderstandingStages.CreateLineSubset(line, words,
                    options?.ReadingDirection ?? PdfReadingDirection.Auto,
                    analysis.ConsumeWork, analysis.CancellationCheck));
                segmentStart = end;
            }
        }
        return result;

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
