namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record PdfBinaryClassificationScore(
    int TruePositive,
    int FalsePositive,
    int FalseNegative,
    double Precision,
    double Recall,
    double F1);

internal sealed record PdfUnderstandingAccuracyObservation(
    int PageCount,
    int ExpectedMarkers,
    int MatchedMarkers,
    double LabelledRegionCharacterErrorRate,
    double PairwiseReadingOrderAccuracy,
    double KendallTau,
    IReadOnlyList<string> ReadingOrderMismatches,
    IReadOnlyDictionary<string, PdfBinaryClassificationScore> Classifications);

public sealed record PdfUnderstandingPerformanceObservation(
    int PageCount,
    int WordCount,
    int RegionCount,
    int SemanticElementCount,
    int TableLikeElementCount);

public sealed record PdfLogicalStructureObservation(
    int PageCount,
    int TextBlockCount,
    int HeadingCount,
    int TableCount,
    int TableCellCount);

internal static class PdfUnderstandingBenchmarkValidation {
    internal static PdfUnderstandingAccuracyObservation Evaluate(
        PdfUnderstandingResult result,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        if (result.Pages.Count != expectedPages.Count) {
            throw new InvalidDataException($"Understanding produced {result.Pages.Count} pages; expected {expectedPages.Count}.");
        }

        int expectedMarkerCount = 0;
        int matchedMarkerCount = 0;
        long correctPairs = 0;
        long totalPairs = 0;
        var expectedSequence = new List<string>();
        var actualSequence = new List<string>();
        var expectedKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var truePositiveKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var predictedKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var readingOrderMismatches = new List<string>();

        for (int pageIndex = 0; pageIndex < expectedPages.Count; pageIndex++) {
            PdfUnderstandingPageResult page = result.Pages[pageIndex];
            PdfUnderstandingBenchmarkExpectation expected = expectedPages[pageIndex];
            if (page.PageNumber != expected.PageNumber) {
                throw new InvalidDataException($"Understanding page {pageIndex + 1} reported source page {page.PageNumber}; expected {expected.PageNumber}.");
            }

            var actualPositions = new Dictionary<string, int>(StringComparer.Ordinal);
            var actualTextByMarker = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int markerIndex = 0; markerIndex < expected.ReadingOrder.Count; markerIndex++) {
                string marker = expected.ReadingOrder[markerIndex];
                expectedSequence.Add(expected.ExpectedRegionText[marker]);
                expectedMarkerCount++;
                int position = FindContainingIndex(page.ReadingOrder.Select(static region => region.Text), marker);
                if (position >= 0) {
                    actualPositions.Add(marker, position);
                    actualTextByMarker.Add(marker, page.ReadingOrder[position].Text);
                    matchedMarkerCount++;
                }
            }

            string[] actualPageSequence = expected.ReadingOrder
                .Where(actualPositions.ContainsKey)
                .OrderBy(marker => actualPositions[marker])
                .ToArray();
            actualSequence.AddRange(actualPageSequence.Select(marker => actualTextByMarker[marker]));
            if (!actualPageSequence.SequenceEqual(expected.ReadingOrder)) {
                readingOrderMismatches.Add(
                    $"Page {page.PageNumber}: expected [{string.Join(", ", expected.ReadingOrder)}], actual [{string.Join(", ", actualPageSequence)}].");
            }

            for (int left = 0; left < expected.ReadingOrder.Count; left++) {
                for (int right = left + 1; right < expected.ReadingOrder.Count; right++) {
                    totalPairs++;
                    if (actualPositions.TryGetValue(expected.ReadingOrder[left], out int leftPosition) &&
                        actualPositions.TryGetValue(expected.ReadingOrder[right], out int rightPosition) &&
                        leftPosition < rightPosition) {
                        correctPairs++;
                    }
                }
            }

            foreach (PdfUnderstandingSemanticElement element in page.Elements) {
                Increment(predictedKinds, element.Kind);
            }

            var matchedElementIndexes = new HashSet<int>();
            foreach (KeyValuePair<string, PdfUnderstandingSemanticKind> pair in expected.SemanticKinds) {
                Increment(expectedKinds, pair.Value);
                int predictedIndex = FindUnmatchedElementIndex(page.Elements, matchedElementIndexes, pair.Key);
                if (predictedIndex >= 0) {
                    matchedElementIndexes.Add(predictedIndex);
                }
                if (predictedIndex >= 0 && page.Elements[predictedIndex].Kind == pair.Value) {
                    Increment(truePositiveKinds, pair.Value);
                }
            }
        }

        string expectedText = string.Join("\n", expectedSequence);
        string actualText = string.Join("\n", actualSequence);
        double characterErrorRate = expectedText.Length == 0
            ? 0D
            : (double)LevenshteinDistance(expectedText, actualText) / expectedText.Length;
        double pairwiseAccuracy = totalPairs == 0 ? 1D : (double)correctPairs / totalPairs;
        double kendallTau = (2D * pairwiseAccuracy) - 1D;
        var classifications = new Dictionary<string, PdfBinaryClassificationScore>(StringComparer.Ordinal);
        foreach (PdfUnderstandingSemanticKind kind in Enum.GetValues<PdfUnderstandingSemanticKind>()) {
            int expected = expectedKinds.TryGetValue(kind, out int expectedCount) ? expectedCount : 0;
            int predicted = predictedKinds.TryGetValue(kind, out int predictedCount) ? predictedCount : 0;
            int truePositive = truePositiveKinds.TryGetValue(kind, out int truePositiveCount) ? truePositiveCount : 0;
            int falsePositive = Math.Max(0, predicted - truePositive);
            int falseNegative = Math.Max(0, expected - truePositive);
            double precision = truePositive + falsePositive == 0 ? 1D : (double)truePositive / (truePositive + falsePositive);
            double recall = truePositive + falseNegative == 0 ? 1D : (double)truePositive / (truePositive + falseNegative);
            double f1 = precision + recall == 0D ? 0D : 2D * precision * recall / (precision + recall);
            classifications.Add(kind.ToString(), new PdfBinaryClassificationScore(
                truePositive,
                falsePositive,
                falseNegative,
                precision,
                recall,
                f1));
        }

        return new PdfUnderstandingAccuracyObservation(
            result.Pages.Count,
            expectedMarkerCount,
            matchedMarkerCount,
            characterErrorRate,
            pairwiseAccuracy,
            kendallTau,
            readingOrderMismatches.AsReadOnly(),
            classifications);
    }

    internal static PdfUnderstandingPerformanceObservation Observe(PdfUnderstandingResult result) => new(
        result.Pages.Count,
        result.Pages.Sum(static page => page.Words.Count),
        result.Pages.Sum(static page => page.Regions.Count),
        result.Pages.Sum(static page => page.Elements.Count),
        result.Pages.Sum(static page => page.Elements.Count(element => element.Kind == PdfUnderstandingSemanticKind.Table)));

    internal static PdfLogicalStructureObservation Observe(PdfLogicalDocument logical) {
        int tableCellCount = 0;
        for (int tableIndex = 0; tableIndex < logical.Tables.Count; tableIndex++) {
            PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(logical.Tables[tableIndex]);
            tableCellCount += data.Rows.Sum(static row => row.Count) + data.Columns.Count;
        }

        return new PdfLogicalStructureObservation(
            logical.Pages.Count,
            logical.TextBlocks.Count,
            logical.Headings.Count,
            logical.Tables.Count,
            tableCellCount);
    }

    internal static PdfBinaryClassificationScore EvaluateTableDetection(
        PdfLogicalDocument logical,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        var predictedTables = new List<(int PageNumber, string Text)>(logical.Tables.Count);
        for (int tableIndex = 0; tableIndex < logical.Tables.Count; tableIndex++) {
            PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(logical.Tables[tableIndex]);
            predictedTables.Add((
                logical.Tables[tableIndex].PageNumber,
                string.Join(" ", data.Columns.Concat(data.Rows.SelectMany(static row => row)))));
        }

        int truePositive = 0;
        var matchedTables = new HashSet<int>();
        for (int pageIndex = 0; pageIndex < expectedPages.Count; pageIndex++) {
            string marker = expectedPages[pageIndex].TableMarker;
            for (int tableIndex = 0; tableIndex < predictedTables.Count; tableIndex++) {
                if (!matchedTables.Contains(tableIndex) &&
                    predictedTables[tableIndex].PageNumber == expectedPages[pageIndex].PageNumber &&
                    predictedTables[tableIndex].Text.Contains(marker, StringComparison.Ordinal)) {
                    matchedTables.Add(tableIndex);
                    truePositive++;
                    break;
                }
            }
        }

        int falsePositive = Math.Max(0, predictedTables.Count - truePositive);
        int falseNegative = Math.Max(0, expectedPages.Count - truePositive);
        double precision = truePositive + falsePositive == 0 ? 0D : (double)truePositive / (truePositive + falsePositive);
        double recall = truePositive + falseNegative == 0 ? 0D : (double)truePositive / (truePositive + falseNegative);
        double f1 = precision + recall == 0D ? 0D : 2D * precision * recall / (precision + recall);
        return new PdfBinaryClassificationScore(truePositive, falsePositive, falseNegative, precision, recall, f1);
    }

    internal static void RequireCompleteLabelCoverage(PdfUnderstandingAccuracyObservation observation) {
        if (observation.MatchedMarkers != observation.ExpectedMarkers) {
            throw new InvalidDataException($"Understanding matched {observation.MatchedMarkers}/{observation.ExpectedMarkers} labelled markers.");
        }
        if (observation.PairwiseReadingOrderAccuracy < 1D) {
            throw new InvalidDataException(
                $"Reading-order accuracy was {observation.PairwiseReadingOrderAccuracy:P2}; expected 100% for the deterministic benchmark corpus. {observation.ReadingOrderMismatches.FirstOrDefault()}");
        }
    }

    internal static void RequireDeterministicSemanticQuality(PdfUnderstandingAccuracyObservation observation) {
        KeyValuePair<string, PdfBinaryClassificationScore> incomplete = observation.Classifications.FirstOrDefault(
            static pair => pair.Value.Precision < 1D || pair.Value.Recall < 1D);
        if (!string.IsNullOrEmpty(incomplete.Key)) {
            throw new InvalidDataException(
                $"{incomplete.Key} classification precision/recall was {incomplete.Value.Precision:P2}/{incomplete.Value.Recall:P2}; expected 100% for the deterministic benchmark corpus.");
        }
    }

    internal static void RequireDeterministicTableQuality(PdfBinaryClassificationScore observation) {
        if (observation.Precision < 1D || observation.Recall < 1D) {
            throw new InvalidDataException(
                $"Logical table detection precision/recall was {observation.Precision:P2}/{observation.Recall:P2}; expected 100% for the deterministic benchmark corpus.");
        }
    }

    private static int FindContainingIndex(IEnumerable<string> values, string marker) {
        int index = 0;
        foreach (string value in values) {
            if (value.Contains(marker, StringComparison.Ordinal)) {
                return index;
            }
            index++;
        }
        return -1;
    }

    private static int FindUnmatchedElementIndex(
        IReadOnlyList<PdfUnderstandingSemanticElement> elements,
        HashSet<int> matchedIndexes,
        string marker) {
        for (int index = 0; index < elements.Count; index++) {
            if (!matchedIndexes.Contains(index) && elements[index].Region.Text.Contains(marker, StringComparison.Ordinal)) {
                return index;
            }
        }
        return -1;
    }

    private static void Increment(Dictionary<PdfUnderstandingSemanticKind, int> counts, PdfUnderstandingSemanticKind kind) {
        counts.TryGetValue(kind, out int count);
        counts[kind] = count + 1;
    }

    private static int LevenshteinDistance(string expected, string actual) {
        if (expected.Length == 0) return actual.Length;
        if (actual.Length == 0) return expected.Length;

        var previous = new int[actual.Length + 1];
        var current = new int[actual.Length + 1];
        for (int index = 0; index <= actual.Length; index++) previous[index] = index;
        for (int expectedIndex = 1; expectedIndex <= expected.Length; expectedIndex++) {
            current[0] = expectedIndex;
            for (int actualIndex = 1; actualIndex <= actual.Length; actualIndex++) {
                int substitution = previous[actualIndex - 1] + (expected[expectedIndex - 1] == actual[actualIndex - 1] ? 0 : 1);
                current[actualIndex] = Math.Min(
                    Math.Min(previous[actualIndex] + 1, current[actualIndex - 1] + 1),
                    substitution);
            }
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }
}
