using System.Collections.ObjectModel;
using System.Threading;

using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>
/// Projects accepted OCR geometry into the canonical PDF understanding pipeline. This type owns
/// only OCR evidence normalization; table, region, reading-order, and semantic decisions remain
/// owned by the shared understanding stages.
/// </summary>
internal static class PdfOcrLogicalDocumentBuilder {
    private const double MinimumVisualRunGapPoints = 18D;

    internal static IReadOnlyList<PdfRecognizedWord> OrderWordsForLogicalReading(
        IReadOnlyList<PdfRecognizedWord> words,
        PdfLogicalPage canonicalPage,
        PdfReadingDirection readingDirection,
        CancellationToken cancellationToken) =>
        OrderWordsForLogicalReadingCore(words, canonicalPage, readingDirection, cancellationToken);

    private static ReadOnlyCollection<PdfRecognizedWord> OrderWordsForLogicalReadingCore(
        IReadOnlyList<PdfRecognizedWord> words,
        PdfLogicalPage canonicalPage,
        PdfReadingDirection readingDirection,
        CancellationToken cancellationToken) {
        var bySequence = words.ToDictionary(static word => word.ProviderSequence);
        var emitted = new HashSet<int>();
        var result = new List<PdfRecognizedWord>(words.Count);
        for (int lineIndex = 0; lineIndex < canonicalPage.Analysis.LogicalProjectionLines.Count; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfUnderstandingLine line = canonicalPage.Analysis.LogicalProjectionLines[lineIndex];
            if (line.SourceKind != PdfLogicalContentSourceKind.Ocr) continue;
            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                int? sequence = line.Words[wordIndex].SourceSequence;
                if (!sequence.HasValue || !emitted.Add(sequence.Value) ||
                    !bySequence.TryGetValue(sequence.Value, out PdfRecognizedWord? word)) continue;
                result.Add(word);
            }
        }

        // Custom structural stages may intentionally omit OCR lines from the logical projection.
        // The searchable layer must still contain every accepted word, so append only the unmatched
        // words using the local provider/geometry ordering contract.
        foreach (PdfRecognizedWord word in BuildLines(words, readingDirection, cancellationToken).SelectMany(static line => line.OrderedWords)) {
            if (emitted.Add(word.ProviderSequence)) result.Add(word);
        }
        return result.AsReadOnly();
    }

    internal static IReadOnlyList<PdfOcrLogicalTextLine> BuildTextLines(
        IReadOnlyList<PdfRecognizedWord> words,
        CancellationToken cancellationToken) =>
        BuildTextLines(words, PdfReadingDirection.Auto, cancellationToken);

    internal static IReadOnlyList<PdfOcrLogicalTextLine> BuildTextLines(
        IReadOnlyList<PdfRecognizedWord> words,
        PdfReadingDirection readingDirection,
        CancellationToken cancellationToken) =>
        BuildLines(words, readingDirection, cancellationToken)
            .Select(static line => new PdfOcrLogicalTextLine(line.Top, line.Left, line.Text))
            .ToArray();

    internal static PdfDocumentReadResult Build(
        PdfReadDocument sourceDocument,
        PdfDocumentReadResult nativeDocument,
        IReadOnlyList<PdfUnderstandingPageResult> nativePageAnalyses,
        IReadOnlyList<PdfOcrPageMergeResult> mergePages,
        PdfTextLayoutOptions layoutOptions,
        PdfUnderstandingPipelineOptions pipelineOptions,
        CancellationToken cancellationToken) {
        if (mergePages.All(static page => page.Words.Count == 0)) return nativeDocument;
        if (nativePageAnalyses.Count != nativeDocument.Pages.Count) {
            throw new ArgumentException(
                "Native page-analysis count must match the logical page count.",
                nameof(nativePageAnalyses));
        }

        var mergesByPage = new Dictionary<int, Queue<PdfOcrPageMergeResult>>();
        for (int mergeIndex = 0; mergeIndex < mergePages.Count; mergeIndex++) {
            PdfOcrPageMergeResult mergePage = mergePages[mergeIndex];
            if (!mergesByPage.TryGetValue(mergePage.PageNumber, out Queue<PdfOcrPageMergeResult>? pageMerges)) {
                pageMerges = new Queue<PdfOcrPageMergeResult>();
                mergesByPage.Add(mergePage.PageNumber, pageMerges);
            }
            pageMerges.Enqueue(mergePage);
        }

        var pipeline = new PdfUnderstandingPipeline(layoutOptions, pipelineOptions);
        var analyses = nativePageAnalyses.ToArray();
        for (int pageIndex = 0; pageIndex < nativeDocument.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalPage nativePage = nativeDocument.Pages[pageIndex];
            if (!mergesByPage.TryGetValue(nativePage.PageNumber, out Queue<PdfOcrPageMergeResult>? pageMerges) ||
                pageMerges.Count == 0) continue;

            PdfOcrPageMergeResult mergePage = pageMerges.Dequeue();
            if (mergePage.Words.Count == 0) continue;
            PdfReadPage sourcePage = sourceDocument.Pages[nativePage.PageNumber - 1];
            OcrArtifacts ocr = BuildArtifacts(nativePage, mergePage.Words, layoutOptions.ReadingDirection, cancellationToken);
            PdfUnderstandingPageResult nativeAnalysis = nativePageAnalyses[pageIndex];
            PdfUnderstandingWord[] combinedWords = nativeAnalysis.Words.Concat(ocr.Words).ToArray();
            PdfUnderstandingLine[] combinedLines = nativeAnalysis.Lines.Concat(ocr.Lines).ToArray();
            analyses[pageIndex] = pipeline.RunPositionedPage(
                sourcePage,
                nativePage.PageNumber,
                nativeAnalysis.DecodedRuns,
                combinedWords,
                combinedLines,
                typeof(PdfOcrLogicalDocumentBuilder),
                cancellationToken);
        }

        int[] pageNumbers = nativeDocument.Pages.Select(static page => page.PageNumber).ToArray();
        IReadOnlyList<PdfUnderstandingPageResult> enrichedAnalyses = nativeDocument.Profile == PdfReadProfile.Structured
            ? PdfDocumentSemanticEnricher.Enrich(
                sourceDocument,
                pageNumbers,
                analyses,
                pipelineOptions.MaxRegionsPerPage,
                pipelineOptions.MaxDocumentWorkUnits,
                cancellationToken)
            : analyses;
        PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(
            sourceDocument,
            layoutOptions,
            pageNumbers,
            enrichedAnalyses,
            nativeDocument.Profile,
            cancellationToken);
        for (int pageIndex = 0; pageIndex < enrichedAnalyses.Count; pageIndex++) {
            enrichedAnalyses[pageIndex].CompleteOperation();
        }
        return result;
    }

    private static OcrArtifacts BuildArtifacts(
        PdfLogicalPage page,
        IReadOnlyList<PdfRecognizedWord> words,
        PdfReadingDirection readingDirection,
        CancellationToken cancellationToken) {
        List<OcrLine> sourceLines = BuildLines(words, readingDirection, cancellationToken);
        var understandingWords = new List<PdfUnderstandingWord>(words.Count);
        var projectedBySource = new Dictionary<PdfRecognizedWord, PdfUnderstandingWord>(words.Count);
        for (int wordIndex = 0; wordIndex < words.Count; wordIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfRecognizedWord sourceWord = words[wordIndex];
            PdfUnderstandingWord projected = ProjectWord(page, sourceWord);
            understandingWords.Add(projected);
            projectedBySource.Add(sourceWord, projected);
        }
        var baselineLines = new List<PdfUnderstandingLine>(sourceLines.Count);
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OcrLine source = sourceLines[lineIndex];
            baselineLines.Add(CreateUnderstandingLine(source, source.Words, projectedBySource, readingDirection));
        }

        var understandingLines = new List<PdfUnderstandingLine>(baselineLines.Count);
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OcrLine source = sourceLines[lineIndex];
            PdfUnderstandingLine baselineLine = baselineLines[lineIndex];
            if (source.LineId is not null) {
                understandingLines.Add(baselineLine);
                continue;
            }

            IReadOnlyList<OcrVisualRun> visualRuns = SplitVisualRuns(source.Words);
            if (visualRuns.Count == 1) {
                understandingLines.Add(baselineLine);
                continue;
            }
            for (int runIndex = 0; runIndex < visualRuns.Count; runIndex++) {
                understandingLines.Add(CreateUnderstandingLine(source, visualRuns[runIndex].Words, projectedBySource, readingDirection));
            }
        }
        return new OcrArtifacts(
            understandingWords.AsReadOnly(),
            understandingLines.AsReadOnly());
    }

    private static PdfUnderstandingLine CreateUnderstandingLine(
        OcrLine source,
        IReadOnlyList<PdfRecognizedWord> words,
        Dictionary<PdfRecognizedWord, PdfUnderstandingWord> projectedBySource,
        PdfReadingDirection readingDirection) {
        PdfRecognizedWord[] providerOrder = words.OrderBy(static word => word.ProviderSequence).ToArray();
        PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
            readingDirection,
            providerOrder.Select(static word => word.Text));
        PdfRecognizedWord[] orderedWords = source.LineId is not null
            ? providerOrder
            : direction == PdfReadingDirection.RightToLeft
                ? words.OrderByDescending(static word => word.X).ThenBy(static word => word.ProviderSequence).ToArray()
                : words.OrderBy(static word => word.X).ThenBy(static word => word.ProviderSequence).ToArray();
        PdfUnderstandingWord[] projectedWords = orderedWords.Select(word => projectedBySource[word]).ToArray();
        double left = orderedWords.Min(static word => word.X);
        double top = orderedWords.Min(static word => word.Y);
        double right = orderedWords.Max(static word => word.X + word.Width);
        double bottom = orderedWords.Max(static word => word.Y + word.Height);
        double confidence = orderedWords.Average(static word => word.Confidence);
        return new PdfUnderstandingLine(
            Array.AsReadOnly(projectedWords),
            JoinWords(orderedWords),
            confidence,
            new[] { new PdfInferenceEvidence(
                source.LineId is null ? "line.ocr-geometry" : "line.ocr-provider-hierarchy",
                source.LineId is null
                    ? "OCR words form a continuous visual run on a shared baseline."
                    : "OCR words share a provider-supplied block, paragraph, and line hierarchy.",
                source.LineId is null ? 0.65D : 0.95D) },
            PdfLogicalContentSourceKind.Ocr,
            source.LineId is null ? null : orderedWords.Min(static word => word.ProviderSequence),
            source.BlockId,
            source.ParagraphId,
            source.LineId,
            new PdfLogicalVisualBounds(left, top, right, bottom));
    }

    private static PdfUnderstandingWord ProjectWord(PdfLogicalPage page, PdfRecognizedWord word) {
        double visualBaseline = word.Y + word.Height;
        PdfPagePoint start = page.MapVisualPointToUserSpace(word.X, visualBaseline);
        PdfPagePoint end = page.MapVisualPointToUserSpace(word.X + word.Width, visualBaseline);
        PdfPagePoint top = page.MapVisualPointToUserSpace(word.X, word.Y);
        double advance = Distance(start, end);
        double fontSize = Math.Max(1D, Distance(start, top));
        double rotation = Math.Atan2(end.Y - start.Y, end.X - start.X) * 180D / Math.PI;
        return new PdfUnderstandingWord(
            word.Text,
            Math.Min(start.X, end.X),
            Math.Max(start.X, end.X),
            start.Y,
            fontSize,
            rotation,
            Array.Empty<PdfTextSpan>(),
            word.Confidence,
            new[] { new PdfInferenceEvidence(
                "word.ocr-bounds",
                "The word geometry was normalized from rendered-page OCR bounds into PDF user space.",
                word.Confidence - 0.5D) },
            advance,
            new PdfLogicalVisualBounds(word.X, word.Y, word.X + word.Width, word.Y + word.Height),
            word.ProviderSequence);
    }

    private static double Distance(PdfPagePoint left, PdfPagePoint right) {
        double x = right.X - left.X;
        double y = right.Y - left.Y;
        return Math.Sqrt((x * x) + (y * y));
    }

    private static List<OcrLine> BuildLines(
        IReadOnlyList<PdfRecognizedWord> words,
        PdfReadingDirection readingDirection,
        CancellationToken cancellationToken) {
        const double centerBucketSize = 4D;
        var lines = new List<OcrLine>();
        foreach (var group in words
            .Where(static word => word.LineId is not null)
            .GroupBy(static word => new { word.BlockId, word.ParagraphId, word.LineId })
            .OrderBy(static group => group.Min(static word => word.ProviderSequence))) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfRecognizedWord[] lineWords = group.OrderBy(static word => word.ProviderSequence).ToArray();
            var line = new OcrLine(
                lineWords[0].ProviderSequence,
                readingDirection,
                lineWords[0].BlockId,
                lineWords[0].ParagraphId,
                lineWords[0].LineId);
            for (int wordIndex = 0; wordIndex < lineWords.Length; wordIndex++) line.Add(lineWords[wordIndex]);
            lines.Add(line);
        }
        int nextLineSequence = lines.Count == 0
            ? 0
            : lines.Max(static item => item.Sequence) + 1;
        var buckets = new Dictionary<long, List<OcrLine>>();
        PdfRecognizedWord[] orderedWords = words
            .Where(static word => word.LineId is null)
            .OrderBy(static word => word.Y)
            .ThenBy(static word => word.X)
            .ToArray();
        for (int wordIndex = 0; wordIndex < orderedWords.Length; wordIndex++) {
            if ((wordIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            PdfRecognizedWord word = orderedWords[wordIndex];
            double center = word.Y + (word.Height / 2D);
            double maximumCenterDistance = Math.Max(2D, word.Height * 0.6D);
            long firstBucket = GetCenterBucket(center - maximumCenterDistance, centerBucketSize);
            long lastBucket = GetCenterBucket(center + maximumCenterDistance, centerBucketSize);
            OcrLine? line = null;
            for (long bucket = firstBucket; bucket <= lastBucket; bucket++) {
                if (!buckets.TryGetValue(bucket, out List<OcrLine>? candidates)) continue;
                for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                    OcrLine candidate = candidates[candidateIndex];
                    if (Math.Abs(candidate.CenterY - center) <= Math.Max(2D, Math.Min(candidate.Height, word.Height) * 0.6D) &&
                        (line is null || candidate.Sequence > line.Sequence)) {
                        line = candidate;
                    }
                }
            }
            if (line is null) {
                line = new OcrLine(nextLineSequence++, readingDirection);
                line.Add(word);
                lines.Add(line);
                AddToBucket(line, buckets, centerBucketSize);
                continue;
            }
            long previousBucket = GetCenterBucket(line.CenterY, centerBucketSize);
            line.Add(word);
            long currentBucket = GetCenterBucket(line.CenterY, centerBucketSize);
            if (previousBucket != currentBucket) {
                buckets[previousBucket].Remove(line);
                AddToBucket(line, buckets, centerBucketSize);
            }
        }
        List<OcrLine> ordered = lines
            .OrderBy(static line => line.Top)
            .ThenBy(static line => line.Left)
            .ToList();
        var knownSlots = new List<int>();
        var knownLines = new List<OcrLine>();
        for (int index = 0; index < ordered.Count; index++) {
            OcrLine line = ordered[index];
            if (line.LineId is null) continue;
            knownSlots.Add(index);
            knownLines.Add(line);
        }
        knownLines.Sort(static (left, right) => left.Sequence.CompareTo(right.Sequence));
        for (int index = 0; index < knownSlots.Count; index++) ordered[knownSlots[index]] = knownLines[index];
        return ordered;
    }

    private static long GetCenterBucket(double center, double bucketSize) =>
        checked((long)Math.Floor(center / bucketSize));

    private static void AddToBucket(
        OcrLine line,
        Dictionary<long, List<OcrLine>> buckets,
        double bucketSize) {
        long bucket = GetCenterBucket(line.CenterY, bucketSize);
        if (!buckets.TryGetValue(bucket, out List<OcrLine>? bucketLines)) {
            bucketLines = new List<OcrLine>();
            buckets.Add(bucket, bucketLines);
        }
        bucketLines.Add(line);
    }

    private static IReadOnlyList<OcrVisualRun> SplitVisualRuns(List<PdfRecognizedWord> words) {
        if (words.Count == 0) return Array.Empty<OcrVisualRun>();
        PdfRecognizedWord[] positioned = words.OrderBy(static word => word.X).ToArray();
        var runs = new List<OcrVisualRun>();
        var current = new List<PdfRecognizedWord> { positioned[0] };
        for (int index = 1; index < positioned.Length; index++) {
            PdfRecognizedWord previous = positioned[index - 1];
            PdfRecognizedWord word = positioned[index];
            double minimumGap = Math.Max(
                MinimumVisualRunGapPoints,
                Math.Min(previous.Height, word.Height) * 1.25D);
            if (word.X - (previous.X + previous.Width) >= minimumGap) {
                runs.Add(OcrVisualRun.From(current));
                current.Clear();
            }
            current.Add(word);
        }
        runs.Add(OcrVisualRun.From(current));
        return runs;
    }

    private sealed class OcrLine {
        private double _confidenceTotal;
        private readonly PdfReadingDirection _readingDirection;

        internal OcrLine(
            int sequence,
            PdfReadingDirection readingDirection,
            string? blockId = null,
            string? paragraphId = null,
            string? lineId = null) {
            Sequence = sequence;
            _readingDirection = readingDirection;
            BlockId = blockId;
            ParagraphId = paragraphId;
            LineId = lineId;
        }

        internal int Sequence { get; private set; }
        internal string? BlockId { get; }
        internal string? ParagraphId { get; }
        internal string? LineId { get; }
        internal List<PdfRecognizedWord> Words { get; } = new();
        internal double Left { get; private set; }
        internal double Top { get; private set; }
        internal double Right { get; private set; }
        internal double Bottom { get; private set; }
        internal double CenterY => (Top + Bottom) / 2D;
        internal double Height => Bottom - Top;
        internal double Confidence => Words.Count == 0 ? 0D : _confidenceTotal / Words.Count;
        internal IEnumerable<PdfRecognizedWord> OrderedWords {
            get {
                if (LineId is not null) return Words.OrderBy(static word => word.ProviderSequence);
                PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
                    _readingDirection,
                    Words.OrderBy(static word => word.ProviderSequence).Select(static word => word.Text));
                return direction == PdfReadingDirection.RightToLeft
                    ? Words.OrderByDescending(static word => word.X).ThenBy(static word => word.ProviderSequence)
                    : Words.OrderBy(static word => word.X).ThenBy(static word => word.ProviderSequence);
            }
        }
        internal string Text => JoinWords(OrderedWords);

        internal void Add(PdfRecognizedWord word) {
            Sequence = Math.Min(Sequence, word.ProviderSequence);
            if (Words.Count == 0) {
                Left = word.X;
                Top = word.Y;
                Right = word.X + word.Width;
                Bottom = word.Y + word.Height;
            } else {
                Left = Math.Min(Left, word.X);
                Top = Math.Min(Top, word.Y);
                Right = Math.Max(Right, word.X + word.Width);
                Bottom = Math.Max(Bottom, word.Y + word.Height);
            }
            Words.Add(word);
            _confidenceTotal += word.Confidence;
        }
    }

    private sealed class OcrVisualRun {
        private OcrVisualRun(IReadOnlyList<PdfRecognizedWord> words) {
            Words = words;
        }

        internal IReadOnlyList<PdfRecognizedWord> Words { get; }

        internal static OcrVisualRun From(IReadOnlyList<PdfRecognizedWord> words) => new(words.ToArray());
    }

    private sealed class OcrArtifacts {
        internal OcrArtifacts(
            IReadOnlyList<PdfUnderstandingWord> words,
            IReadOnlyList<PdfUnderstandingLine> lines) {
            Words = words;
            Lines = lines;
        }

        internal IReadOnlyList<PdfUnderstandingWord> Words { get; }
        internal IReadOnlyList<PdfUnderstandingLine> Lines { get; }
    }

    private static string JoinWords(IEnumerable<PdfRecognizedWord> source) {
        PdfRecognizedWord[] words = source as PdfRecognizedWord[] ?? source.ToArray();
        var builder = new System.Text.StringBuilder();
        for (int index = 0; index < words.Length; index++) {
            string text = words[index].Text;
            bool visuallyAdjacent = index > 0 && AreVisuallyAdjacent(words[index - 1], words[index]);
            if (builder.Length > 0 && !visuallyAdjacent) builder.Append(' ');
            builder.Append(text);
        }
        return builder.ToString();
    }

    private static bool AreVisuallyAdjacent(PdfRecognizedWord left, PdfRecognizedWord right) {
        double leftEdge = left.X;
        double leftEnd = left.X + left.Width;
        double rightEdge = right.X;
        double rightEnd = right.X + right.Width;
        double gap = rightEdge >= leftEnd
            ? rightEdge - leftEnd
            : leftEdge >= rightEnd
                ? leftEdge - rightEnd
                : 0D;
        return gap <= Math.Max(0.5D, Math.Min(left.Height, right.Height) * 0.12D);
    }

}

internal readonly struct PdfOcrLogicalTextLine {
    internal PdfOcrLogicalTextLine(double top, double left, string text) {
        Top = top;
        Left = left;
        Text = text;
    }

    internal double Top { get; }
    internal double Left { get; }
    internal string Text { get; }
}
