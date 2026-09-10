using System.Threading;

namespace OfficeIMO.Pdf.Ocr;

internal static partial class PdfOcrLogicalDocumentBuilder {
    private static long ApplyRecognitionFrameOrder(
        PdfUnderstandingPipeline pipeline, PdfReadPage sourcePage, PdfLogicalPage nativePage,
        PdfOcrPageMergeResult merge, OcrArtifacts artifacts, CancellationToken token) {
        if (!merge.RecognitionWidth.HasValue || !merge.RecognitionHeight.HasValue || artifacts.Lines.Count < 2 ||
            artifacts.Lines.All(static line => line.SourceSequence.HasValue)) return 0;

        double width = merge.RecognitionWidth.Value;
        double height = merge.RecognitionHeight.Value;
        var normalizedBySequence = new Dictionary<int, PdfUnderstandingWord>(merge.Words.Count);
        foreach (PdfRecognizedWord word in merge.Words) {
            token.ThrowIfCancellationRequested();
            PdfLogicalVisualBounds bounds = word.ReadingBounds;
            normalizedBySequence.Add(word.ProviderSequence, new PdfUnderstandingWord(
                word.Text, bounds.Left, bounds.Right, height - bounds.Bottom, Math.Max(1D, bounds.Height),
                0, Array.Empty<PdfTextSpan>(), word.Confidence, advance: bounds.Width,
                visualBounds: bounds, sourceSequence: word.ProviderSequence));
        }

        var normalizedLines = new List<PdfUnderstandingLine>(artifacts.Lines.Count);
        foreach (PdfUnderstandingLine line in artifacts.Lines) {
            token.ThrowIfCancellationRequested();
            PdfUnderstandingWord[] words = line.Words.Select(word => normalizedBySequence[word.SourceSequence!.Value]).ToArray();
            var bounds = new PdfLogicalVisualBounds(
                words.Min(static word => word.VisualBounds!.Left), words.Min(static word => word.VisualBounds!.Top),
                words.Max(static word => word.VisualBounds!.Right), words.Max(static word => word.VisualBounds!.Bottom));
            normalizedLines.Add(new PdfUnderstandingLine(words, line.Text, line.Confidence, line.Evidence,
                line.SourceKind, line.SourceSequence, line.BlockId, line.ParagraphId, line.LineId, bounds));
        }

        IReadOnlyList<PdfUnderstandingLine> ordered = pipeline.InferPositionedReadingOrder(
            sourcePage, nativePage.PageNumber, normalizedBySequence.Values.ToArray(), normalizedLines,
            width, height, token, out long workUnits);
        var ranks = new Dictionary<int, int>(merge.Words.Count);
        foreach (PdfUnderstandingLine line in ordered) {
            token.ThrowIfCancellationRequested();
            foreach (PdfUnderstandingWord word in line.Words) {
                if (word.SourceSequence.HasValue && !ranks.ContainsKey(word.SourceSequence.Value))
                    ranks.Add(word.SourceSequence.Value, ranks.Count);
            }
        }
        var projected = new List<PdfUnderstandingLine>(artifacts.Lines.Count);
        foreach (PdfUnderstandingLine line in artifacts.Lines) {
            token.ThrowIfCancellationRequested();
            int? rank = null;
            foreach (PdfUnderstandingWord word in line.Words) {
                if (word.SourceSequence.HasValue && ranks.TryGetValue(word.SourceSequence.Value, out int candidate))
                    rank = !rank.HasValue ? candidate : Math.Min(rank.Value, candidate);
            }
            // Keep original-page bounds and quadrilaterals; only the inferred reading rank travels back.
            projected.Add(new PdfUnderstandingLine(line.Words, line.Text, line.Confidence, line.Evidence,
                line.SourceKind, rank ?? line.SourceSequence, line.BlockId, line.ParagraphId, line.LineId, line.VisualBounds));
        }
        artifacts.Lines = projected.AsReadOnly();
        return workUnits;
    }
}