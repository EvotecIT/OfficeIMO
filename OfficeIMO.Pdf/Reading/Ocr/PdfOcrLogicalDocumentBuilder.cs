using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfOcrLogicalDocumentBuilder {
    internal static IReadOnlyList<PdfRecognizedWord> OrderWordsForLogicalReading(
        IReadOnlyList<PdfRecognizedWord> words,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfRecognizedWord[] ordered = words
            .OrderBy(static word => word.ProviderSequence)
            .ToArray();
        cancellationToken.ThrowIfCancellationRequested();
        return ordered;
    }

    internal static IReadOnlyList<PdfOcrLogicalTextLine> BuildTextLines(
        IReadOnlyList<PdfRecognizedWord> words,
        CancellationToken cancellationToken) =>
        BuildLines(words, cancellationToken)
            .Select(static line => new PdfOcrLogicalTextLine(line.Top, line.Left, line.Text))
            .ToArray();

    internal static PdfDocumentReadResult Build(
        PdfDocumentReadResult nativeDocument,
        IReadOnlyList<PdfOcrPageMergeResult> mergePages,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        if (!options.BuildEnrichedLogicalDocument || mergePages.All(static page => page.Words.Count == 0)) {
            return nativeDocument;
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
        var pages = new List<PdfLogicalPage>(nativeDocument.Pages.Count);
        for (int pageIndex = 0; pageIndex < nativeDocument.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalPage nativePage = nativeDocument.Pages[pageIndex];
            if (!mergesByPage.TryGetValue(nativePage.PageNumber, out Queue<PdfOcrPageMergeResult>? pageMerges) ||
                pageMerges.Count == 0) {
                pages.Add(nativePage);
                continue;
            }

            PdfOcrPageMergeResult mergePage = pageMerges.Dequeue();
            pages.Add(mergePage.Words.Count == 0
                ? nativePage
                : EnrichPage(nativePage, mergePage.Words, options, cancellationToken));
        }
        return nativeDocument.WithPages(pages.AsReadOnly());
    }

    private static PdfLogicalPage EnrichPage(
        PdfLogicalPage page,
        IReadOnlyList<PdfRecognizedWord> words,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        List<OcrLine> sourceLines = BuildLines(words, cancellationToken);
        HashSet<int> tableLineIndexes = new HashSet<int>();
        IReadOnlyList<PdfUnderstandingTableCandidate> tableCandidates = options.DetectAlignedTables
            ? DetectTables(page, sourceLines, tableLineIndexes, options, cancellationToken)
            : Array.Empty<PdfUnderstandingTableCandidate>();
        List<(OcrLine Line, bool IsTableLine)> lines = BuildSemanticLines(
            sourceLines,
            tableLineIndexes,
            options.MinimumTableColumnGapPoints,
            cancellationToken);
        double medianHeight = Median(lines.Select(static item => item.Line.Height));
        var textBlocks = new List<PdfLogicalTextBlock>(lines.Count);
        var headings = new List<PdfLogicalHeading>();
        var listItems = new List<PdfLogicalListItem>();
        var paragraphLines = new List<PdfLogicalTextBlock>();
        var paragraphs = new List<PdfLogicalParagraph>();

        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            if ((lineIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            (OcrLine line, bool isTableLine) = lines[lineIndex];
            string marker = string.Empty;
            string listText = line.Text;
            bool isList = !isTableLine && TryParseList(line.Text, out marker, out listText);
            bool isHeading = !isTableLine && !isList && line.Height >= medianHeight * 1.35D && line.Text.Length <= 160;
            PdfLogicalElementKind kind = isHeading
                ? PdfLogicalElementKind.Heading
                : isList
                    ? PdfLogicalElementKind.ListItem
                    : PdfLogicalElementKind.TextBlock;
            PdfVisualBounds userBounds = page.TransformVisualBoundsToUser(line.Left, line.Top, line.Right, line.Bottom);
            var visualBounds = new PdfLogicalVisualBounds(line.Left, line.Top, line.Right, line.Bottom);
            var block = new PdfLogicalTextBlock(
                page.PageNumber,
                kind,
                line.Text,
                userBounds.Left,
                userBounds.Right,
                userBounds.Top,
                Math.Max(1D, line.Height),
                Array.Empty<PdfTextSpan>(),
                PdfLogicalContentSourceKind.Ocr,
                line.Confidence,
                visualBounds);
            textBlocks.Add(block);

            if (isHeading) {
                FlushParagraph();
                headings.Add(new PdfLogicalHeading(page.PageNumber, HeadingLevel(line.Height, medianHeight), line.Text, line.Height, block));
            } else if (isList) {
                FlushParagraph();
                listItems.Add(new PdfLogicalListItem(page.PageNumber, 1, marker, listText, block));
            } else if (!isTableLine) {
                if (paragraphLines.Count > 0 && !CanContinueParagraph(paragraphLines[paragraphLines.Count - 1], block, medianHeight)) {
                    FlushParagraph();
                }
                paragraphLines.Add(block);
            }
        }
        FlushParagraph();

        return page.WithOcrContent(
            textBlocks.AsReadOnly(),
            headings.AsReadOnly(),
            paragraphs.AsReadOnly(),
            listItems.AsReadOnly(),
            tableCandidates);

        void FlushParagraph() {
            if (paragraphLines.Count == 0) return;
            paragraphs.Add(PdfLogicalParagraph.FromOcr(page.PageNumber, paragraphLines.ToArray()));
            paragraphLines.Clear();
        }
    }

    private static List<OcrLine> BuildLines(
        IReadOnlyList<PdfRecognizedWord> words,
        CancellationToken cancellationToken) {
        const double centerBucketSize = 4D;
        var lines = new List<OcrLine>();
        var buckets = new Dictionary<long, List<OcrLine>>();
        PdfRecognizedWord[] orderedWords = words.OrderBy(static word => word.Y).ThenBy(static word => word.X).ToArray();
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
                line = new OcrLine(lines.Count);
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
        foreach (OcrLine line in lines) line.Words.Sort(static (left, right) => left.X.CompareTo(right.X));
        return lines.OrderBy(static line => line.Top).ThenBy(static line => line.Left).ToList();
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

    private static List<(OcrLine Line, bool IsTableLine)> BuildSemanticLines(
        IReadOnlyList<OcrLine> sourceLines,
        HashSet<int> tableLineIndexes,
        double minimumGap,
        CancellationToken cancellationToken) {
        var result = new List<(OcrLine Line, bool IsTableLine)>();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            if ((lineIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            OcrLine line = sourceLines[lineIndex];
            bool isTableLine = tableLineIndexes.Contains(lineIndex);
            IReadOnlyList<OcrCell> cells = isTableLine
                ? Array.Empty<OcrCell>()
                : SplitCells(line.Words, minimumGap);
            if (cells.Count < 2) {
                result.Add((line, isTableLine));
                continue;
            }
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                result.Add((OcrLine.FromWords(cells[cellIndex].Words, result.Count), false));
            }
        }
        return result;
    }

    private static IReadOnlyList<PdfUnderstandingTableCandidate> DetectTables(
        PdfLogicalPage page,
        IReadOnlyList<OcrLine> lines,
        HashSet<int> tableLineIndexes,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        var candidates = new List<(int Index, IReadOnlyList<OcrCell> Cells)>();
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            if ((lineIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OcrCell> cells = SplitCells(lines[lineIndex].Words, options.MinimumTableColumnGapPoints);
            if (cells.Count >= 2) candidates.Add((lineIndex, cells));
        }
        if (candidates.Count < options.MinimumAlignedTableRows) return Array.Empty<PdfUnderstandingTableCandidate>();

        var groups = new List<List<(int Index, IReadOnlyList<OcrCell> Cells)>>();
        foreach ((int index, IReadOnlyList<OcrCell> cells) in candidates) {
            cancellationToken.ThrowIfCancellationRequested();
            List<(int Index, IReadOnlyList<OcrCell> Cells)>? group = groups.LastOrDefault();
            if (group is null || index != group[group.Count - 1].Index + 1 || !ColumnsAlign(group[0].Cells, cells, options.TableColumnTolerancePoints)) {
                group = new List<(int Index, IReadOnlyList<OcrCell> Cells)>();
                groups.Add(group);
            }
            group.Add((index, cells));
        }

        var result = new List<PdfUnderstandingTableCandidate>();
        foreach (List<(int Index, IReadOnlyList<OcrCell> Cells)> group in groups.Where(group => group.Count >= options.MinimumAlignedTableRows)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HasConservativeTableEvidence(group, lines)) continue;
            if (result.Count >= options.MaxInferredTablesPerPage) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxInferredTablesPerPage, result.Count + 1L);
            }
            int columnCount = group[0].Cells.Count;
            var visualColumnBounds = new List<(double From, double To)>(columnCount);
            for (int column = 0; column < columnCount; column++) {
                visualColumnBounds.Add((group.Min(row => row.Cells[column].Left), group.Max(row => row.Cells[column].Right)));
            }
            var rows = group.Select(row => (IReadOnlyList<string>)row.Cells.Select(static cell => cell.Text).ToArray()).ToArray();
            double top = group.Min(row => lines[row.Index].Top);
            double bottom = group.Max(row => lines[row.Index].Bottom);
            double left = visualColumnBounds.Min(static column => column.From);
            double right = visualColumnBounds.Max(static column => column.To);
            double confidence = PdfInference.Clamp(group.Average(row => lines[row.Index].Confidence));
            result.Add(PdfUnderstandingTableCandidate.FromOcr(
                "OcrAlignedColumns",
                top,
                bottom,
                new PdfLogicalVisualBounds(left, top, right, bottom),
                visualColumnBounds,
                rows,
                confidence,
                new[] {
                    new PdfInferenceEvidence(
                        "table.ocr-aligned-geometry",
                        "Accepted OCR words form repeated columns with compact row rhythm.",
                        0.85D)
                }));
            foreach ((int index, IReadOnlyList<OcrCell> _) in group) tableLineIndexes.Add(index);
        }
        return result.AsReadOnly();
    }

    private static bool HasConservativeTableEvidence(
        IReadOnlyList<(int Index, IReadOnlyList<OcrCell> Cells)> rows,
        IReadOnlyList<OcrLine> lines) {
        if (rows.Count < 3 || rows[0].Cells.Count < 2) return false;
        if (!HasCompactCellStructure(rows)) return false;
        double[] verticalSteps = new double[rows.Count - 1];
        for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++) {
            verticalSteps[rowIndex - 1] = lines[rows[rowIndex].Index].CenterY - lines[rows[rowIndex - 1].Index].CenterY;
            if (verticalSteps[rowIndex - 1] <= 0D) return false;
        }
        double medianStep = Median(verticalSteps);
        double medianHeight = Median(rows.Select(row => lines[row.Index].Height));
        if (medianHeight <= 0D || medianStep > Math.Max(24D, medianHeight * 3D)) return false;
        double smallestStep = verticalSteps.Min();
        double largestStep = verticalSteps.Max();
        return smallestStep > 0D && largestStep <= smallestStep * 1.75D;
    }

    private static bool HasCompactCellStructure(
        IReadOnlyList<(int Index, IReadOnlyList<OcrCell> Cells)> rows) {
        long cellCount = 0L;
        long wordCount = 0L;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IReadOnlyList<OcrCell> cells = rows[rowIndex].Cells;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                int words = cells[cellIndex].Words.Count;
                if (words == 0 || words > 4) return false;
                cellCount++;
                wordCount += words;
            }
        }
        return cellCount > 0L && wordCount <= cellCount * 2L;
    }

    private static IReadOnlyList<OcrCell> SplitCells(List<PdfRecognizedWord> words, double minimumGap) {
        if (words.Count == 0) return Array.Empty<OcrCell>();
        var cells = new List<OcrCell>();
        var current = new List<PdfRecognizedWord> { words[0] };
        for (int index = 1; index < words.Count; index++) {
            PdfRecognizedWord previous = words[index - 1];
            PdfRecognizedWord word = words[index];
            if (word.X - (previous.X + previous.Width) >= minimumGap) {
                cells.Add(OcrCell.From(current));
                current.Clear();
            }
            current.Add(word);
        }
        cells.Add(OcrCell.From(current));
        return cells;
    }

    private static bool ColumnsAlign(IReadOnlyList<OcrCell> expected, IReadOnlyList<OcrCell> actual, double tolerance) {
        if (expected.Count != actual.Count) return false;
        for (int index = 0; index < expected.Count; index++) {
            if (Math.Abs(expected[index].Left - actual[index].Left) > tolerance) return false;
        }
        return true;
    }

    private static bool TryParseList(string text, out string marker, out string body) {
        return ContentStructureExtractor.TryParseListItemText(text, out marker, out body, out _);
    }

    private static int HeadingLevel(double height, double medianHeight) =>
        height >= medianHeight * 1.8D ? 1 : height >= medianHeight * 1.55D ? 2 : 3;

    private static bool CanContinueParagraph(PdfLogicalTextBlock previous, PdfLogicalTextBlock current, double medianHeight) {
        PdfLogicalVisualBounds? previousBounds = previous.VisualBounds;
        PdfLogicalVisualBounds? currentBounds = current.VisualBounds;
        if (previousBounds is null || currentBounds is null) return false;
        double verticalGap = currentBounds.Top - previousBounds.Bottom;
        return verticalGap <= Math.Max(4D, medianHeight * 1.1D) &&
            Math.Abs(currentBounds.Left - previousBounds.Left) <= Math.Max(18D, medianHeight * 2D);
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(static value => value).ToArray();
        if (ordered.Length == 0) return 1D;
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }

    private sealed class OcrLine {
        private double _confidenceTotal;
        internal OcrLine(int sequence) { Sequence = sequence; }
        internal int Sequence { get; }
        internal List<PdfRecognizedWord> Words { get; } = new List<PdfRecognizedWord>();
        internal double Left { get; private set; }
        internal double Top { get; private set; }
        internal double Right { get; private set; }
        internal double Bottom { get; private set; }
        internal double CenterY => (Top + Bottom) / 2D;
        internal double Height => Bottom - Top;
        internal double Confidence => Words.Count == 0 ? 0D : _confidenceTotal / Words.Count;
        internal string Text => JoinWords(Words.OrderBy(static word => word.ProviderSequence));
        internal void Add(PdfRecognizedWord word) {
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

        internal static OcrLine FromWords(IReadOnlyList<PdfRecognizedWord> words, int sequence) {
            var line = new OcrLine(sequence);
            for (int index = 0; index < words.Count; index++) line.Add(words[index]);
            return line;
        }
    }

    private sealed class OcrCell {
        private OcrCell(double left, double right, string text, IReadOnlyList<PdfRecognizedWord> words) {
            Left = left;
            Right = right;
            Text = text;
            Words = words;
        }
        internal double Left { get; }
        internal double Right { get; }
        internal string Text { get; }
        internal IReadOnlyList<PdfRecognizedWord> Words { get; }
        internal static OcrCell From(IReadOnlyList<PdfRecognizedWord> words) =>
            new OcrCell(
                words.Min(static word => word.X),
                words.Max(static word => word.X + word.Width),
                JoinWords(words.OrderBy(static word => word.ProviderSequence)),
                words.ToArray());
    }

    private static string JoinWords(IEnumerable<PdfRecognizedWord> source) {
        PdfRecognizedWord[] words = source as PdfRecognizedWord[] ?? source.ToArray();
        var builder = new System.Text.StringBuilder();
        for (int index = 0; index < words.Length; index++) {
            string text = words[index].Text;
            bool attach = index > 0 && text.Length > 0 && IsTrailingPunctuation(text[0]);
            if (builder.Length > 0 && !attach) builder.Append(' ');
            builder.Append(text);
        }
        return builder.ToString();
    }

    private static bool IsTrailingPunctuation(char character) =>
        character == ',' || character == '.' || character == ';' || character == ':' ||
        character == '!' || character == '?' || character == ')' || character == ']' || character == '}';
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
