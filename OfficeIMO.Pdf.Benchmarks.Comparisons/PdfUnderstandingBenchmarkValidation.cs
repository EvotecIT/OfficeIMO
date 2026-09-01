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

internal sealed record PdfHeadingAccuracyObservation(
    PdfBinaryClassificationScore Detection,
    PdfBinaryClassificationScore ExactLevel,
    IReadOnlyList<string> Mismatches);

internal sealed record PdfSemanticCorrectnessObservation(
    PdfUnderstandingAccuracyObservation Regions,
    PdfHeadingAccuracyObservation Headings,
    PdfBinaryClassificationScore TableDetection,
    PdfBinaryClassificationScore CellAdjacency,
    PdfBinaryClassificationScore ContinuationPairs);

public sealed record PdfStructuredReadObservation(
    int PageCount,
    int WordCount,
    int RegionCount,
    int SemanticElementCount,
    int TextBlockCount,
    int HeadingCount,
    int TableCount,
    int TableCellCount,
    int CrossPageTableGroupCount);

internal static class PdfUnderstandingBenchmarkValidation {
    internal static PdfSemanticCorrectnessObservation Evaluate(
        PdfDocumentReadResult document,
        PdfUnderstandingBenchmarkCorpus corpus) {
        PdfUnderstandingAccuracyObservation regions = EvaluateRegions(document, corpus.Pages);
        PdfHeadingAccuracyObservation headings = EvaluateHeadings(document, corpus.Pages);
        PdfBinaryClassificationScore tables = EvaluateTableDetection(document, corpus.Pages);
        PdfBinaryClassificationScore adjacency = EvaluateCellAdjacency(document, corpus.Pages);

        PdfDocument continuationSource = PdfDocument.Load(corpus.ContinuationPdf);
        PdfDocumentReadResult continuationDocument = continuationSource.Read(new PdfReadOptions {
            Profile = PdfReadProfile.Structured
        });
        PdfBinaryClassificationScore continuationPairs = EvaluateContinuationPairs(
            continuationDocument,
            corpus.ExpectedContinuationPairs);

        return new PdfSemanticCorrectnessObservation(regions, headings, tables, adjacency, continuationPairs);
    }

    internal static PdfUnderstandingAccuracyObservation EvaluateRegions(
        PdfDocumentReadResult document,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        if (document.Pages.Count != expectedPages.Count) {
            throw new InvalidDataException($"Structured read produced {document.Pages.Count} pages; expected {expectedPages.Count}.");
        }

        int expectedMarkerCount = 0;
        int matchedMarkerCount = 0;
        long correctPairs = 0;
        long totalPairs = 0;
        long labelledRegionCharacterErrors = 0;
        long labelledRegionCharacterCount = 0;
        var expectedKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var truePositiveKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var predictedKinds = new Dictionary<PdfUnderstandingSemanticKind, int>();
        var readingOrderMismatches = new List<string>();

        for (int pageIndex = 0; pageIndex < expectedPages.Count; pageIndex++) {
            PdfUnderstandingPageResult page = document.Pages[pageIndex].Analysis;
            PdfUnderstandingBenchmarkExpectation expected = expectedPages[pageIndex];
            if (page.PageNumber != expected.PageNumber) {
                throw new InvalidDataException($"Structured page {pageIndex + 1} reported source page {page.PageNumber}; expected {expected.PageNumber}.");
            }

            var actualPositions = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int markerIndex = 0; markerIndex < expected.ReadingOrder.Count; markerIndex++) {
                string marker = expected.ReadingOrder[markerIndex];
                string expectedRegionText = expected.ExpectedRegionText[marker];
                labelledRegionCharacterCount += expectedRegionText.Length;
                expectedMarkerCount++;
                int position = FindContainingIndex(page.ReadingOrder.Select(static region => region.Text), marker);
                if (position >= 0) {
                    actualPositions.Add(marker, position);
                    labelledRegionCharacterErrors += LevenshteinDistance(expectedRegionText, page.ReadingOrder[position].Text);
                    matchedMarkerCount++;
                } else {
                    labelledRegionCharacterErrors += expectedRegionText.Length;
                }
            }

            string[] actualPageSequence = expected.ReadingOrder
                .Where(actualPositions.ContainsKey)
                .OrderBy(marker => actualPositions[marker])
                .ToArray();
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
                if (predictedIndex >= 0) matchedElementIndexes.Add(predictedIndex);
                if (predictedIndex >= 0 && page.Elements[predictedIndex].Kind == pair.Value) {
                    Increment(truePositiveKinds, pair.Value);
                }
            }
        }

        double characterErrorRate = labelledRegionCharacterCount == 0
            ? 0D
            : (double)labelledRegionCharacterErrors / labelledRegionCharacterCount;
        double pairwiseAccuracy = totalPairs == 0 ? 1D : (double)correctPairs / totalPairs;
        double kendallTau = (2D * pairwiseAccuracy) - 1D;
        var classifications = new Dictionary<string, PdfBinaryClassificationScore>(StringComparer.Ordinal);
        foreach (PdfUnderstandingSemanticKind kind in Enum.GetValues<PdfUnderstandingSemanticKind>()) {
            int expected = expectedKinds.TryGetValue(kind, out int expectedCount) ? expectedCount : 0;
            int predicted = predictedKinds.TryGetValue(kind, out int predictedCount) ? predictedCount : 0;
            int truePositive = truePositiveKinds.TryGetValue(kind, out int truePositiveCount) ? truePositiveCount : 0;
            classifications.Add(kind.ToString(), Score(truePositive, predicted - truePositive, expected - truePositive));
        }

        return new PdfUnderstandingAccuracyObservation(
            document.Pages.Count,
            expectedMarkerCount,
            matchedMarkerCount,
            characterErrorRate,
            pairwiseAccuracy,
            kendallTau,
            readingOrderMismatches.AsReadOnly(),
            classifications);
    }

    internal static PdfHeadingAccuracyObservation EvaluateHeadings(
        PdfDocumentReadResult document,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        int expectedCount = expectedPages.Sum(static page => page.HeadingLevels.Count);
        int detected = 0;
        int exactLevel = 0;
        var matched = new HashSet<PdfLogicalHeading>();
        var mismatches = new List<string>();
        foreach (PdfUnderstandingBenchmarkExpectation expectedPage in expectedPages) {
            PdfLogicalPage page = document.Pages.Single(candidate => candidate.PageNumber == expectedPage.PageNumber);
            foreach (KeyValuePair<string, int> expectation in expectedPage.HeadingLevels) {
                PdfLogicalHeading? heading = page.Headings.FirstOrDefault(candidate =>
                    !matched.Contains(candidate) && candidate.Text.Contains(expectation.Key, StringComparison.Ordinal));
                if (heading is null) {
                    mismatches.Add($"Page {expectedPage.PageNumber}: missing heading containing {expectation.Key}.");
                    continue;
                }
                matched.Add(heading);
                detected++;
                if (heading.Level == expectation.Value) {
                    exactLevel++;
                } else {
                    mismatches.Add(
                        $"Page {expectedPage.PageNumber}: expected level {expectation.Value}, observed {heading.Level} for {heading.Text}; evidence={string.Join(',', heading.Evidence.Select(static evidence => evidence.Code))}.");
                }
            }
        }

        int predictedCount = document.Headings.Count;
        return new PdfHeadingAccuracyObservation(
            Score(detected, predictedCount - detected, expectedCount - detected),
            Score(exactLevel, predictedCount - exactLevel, expectedCount - exactLevel),
            mismatches.AsReadOnly());
    }

    internal static PdfBinaryClassificationScore EvaluateTableDetection(
        PdfDocumentReadResult document,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        var predictedTables = ReadTables(document);
        int truePositive = 0;
        var matchedTables = new HashSet<int>();
        for (int pageIndex = 0; pageIndex < expectedPages.Count; pageIndex++) {
            PdfUnderstandingBenchmarkExpectation expected = expectedPages[pageIndex];
            for (int tableIndex = 0; tableIndex < predictedTables.Count; tableIndex++) {
                if (!matchedTables.Contains(tableIndex) &&
                    predictedTables[tableIndex].PageNumber == expected.PageNumber &&
                    predictedTables[tableIndex].Cells.Contains(expected.TableMarker, StringComparer.Ordinal) &&
                    predictedTables[tableIndex].Cells.SequenceEqual(expected.ExpectedTableCells, StringComparer.Ordinal)) {
                    matchedTables.Add(tableIndex);
                    truePositive++;
                    break;
                }
            }
        }

        return Score(truePositive, predictedTables.Count - truePositive, expectedPages.Count - truePositive);
    }

    internal static PdfBinaryClassificationScore EvaluateCellAdjacency(
        PdfDocumentReadResult document,
        IReadOnlyList<PdfUnderstandingBenchmarkExpectation> expectedPages) {
        var expectedEdges = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (IGrouping<int, PdfUnderstandingBenchmarkExpectation> pageGroup in expectedPages.GroupBy(static page => page.PageNumber)) {
            int tableIndex = 0;
            foreach (PdfUnderstandingBenchmarkExpectation page in pageGroup) {
                AddAdjacencyEdges(expectedEdges, page.PageNumber, tableIndex++, page.ExpectedTableRows);
            }
        }

        var predictedEdges = new Dictionary<string, int>(StringComparer.Ordinal);
        var tableIndexesByPage = new Dictionary<int, int>();
        foreach ((int pageNumber, _, IReadOnlyList<IReadOnlyList<string>> rows) in ReadTables(document)) {
            tableIndexesByPage.TryGetValue(pageNumber, out int tableIndex);
            AddAdjacencyEdges(predictedEdges, pageNumber, tableIndex, rows);
            tableIndexesByPage[pageNumber] = tableIndex + 1;
        }

        int expectedCount = expectedEdges.Values.Sum();
        int predictedCount = predictedEdges.Values.Sum();
        int truePositive = predictedEdges.Sum(edge =>
            expectedEdges.TryGetValue(edge.Key, out int expectedOccurrences)
                ? Math.Min(edge.Value, expectedOccurrences)
                : 0);
        return Score(truePositive, predictedCount - truePositive, expectedCount - truePositive);
    }

    internal static PdfBinaryClassificationScore EvaluateContinuationPairs(
        PdfDocumentReadResult document,
        IReadOnlyList<(int PreviousPage, int CurrentPage)> expectedPairs) {
        var predicted = new HashSet<(int PreviousPage, int CurrentPage)>();
        foreach (PdfLogicalTableContinuationGroup group in document.GetTableContinuationGroups()) {
            for (int index = 1; index < group.Segments.Count; index++) {
                predicted.Add((group.Segments[index - 1].PageNumber, group.Segments[index].PageNumber));
            }
        }

        var expected = expectedPairs.ToHashSet();
        int truePositive = predicted.Count(expected.Contains);
        return Score(truePositive, predicted.Count - truePositive, expected.Count - truePositive);
    }

    internal static PdfStructuredReadObservation Observe(PdfDocumentReadResult document) {
        int tableCellCount = 0;
        for (int tableIndex = 0; tableIndex < document.Tables.Count; tableIndex++) {
            PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(document.Tables[tableIndex]);
            tableCellCount += data.Rows.Sum(static row => row.Count) + data.Columns.Count;
        }

        return new PdfStructuredReadObservation(
            document.Pages.Count,
            document.Pages.Sum(static page => page.Analysis.Words.Count),
            document.Pages.Sum(static page => page.Analysis.Regions.Count),
            document.Pages.Sum(static page => page.Analysis.Elements.Count),
            document.TextBlocks.Count,
            document.Headings.Count,
            document.Tables.Count,
            tableCellCount,
            document.GetTableContinuationGroups().Count(static group => group.SpansPages));
    }

    internal static void RequireDeterministicQuality(PdfSemanticCorrectnessObservation observation) {
        PdfUnderstandingAccuracyObservation regions = observation.Regions;
        if (regions.MatchedMarkers != regions.ExpectedMarkers) {
            throw new InvalidDataException($"Structured read matched {regions.MatchedMarkers}/{regions.ExpectedMarkers} labelled markers.");
        }
        if (regions.PairwiseReadingOrderAccuracy < 1D) {
            throw new InvalidDataException(
                $"Reading-order accuracy was {regions.PairwiseReadingOrderAccuracy:P2}; expected 100% for the deterministic corpus. {regions.ReadingOrderMismatches.FirstOrDefault()}");
        }
        if (regions.LabelledRegionCharacterErrorRate != 0D) {
            throw new InvalidDataException(
                $"Labelled-region character error rate was {regions.LabelledRegionCharacterErrorRate:P4}; expected 0% for the deterministic corpus.");
        }

        KeyValuePair<string, PdfBinaryClassificationScore> incomplete = regions.Classifications.FirstOrDefault(
            static pair => pair.Value.Precision < 1D || pair.Value.Recall < 1D);
        if (!string.IsNullOrEmpty(incomplete.Key)) {
            throw new InvalidDataException(
                $"{incomplete.Key} classification precision/recall was {incomplete.Value.Precision:P2}/{incomplete.Value.Recall:P2}; expected 100% for the deterministic corpus.");
        }

        RequirePerfect("heading detection", observation.Headings.Detection);
        if (observation.Headings.ExactLevel.Precision < 1D || observation.Headings.ExactLevel.Recall < 1D) {
            throw new InvalidDataException(
                $"Deterministic heading level precision/recall was {observation.Headings.ExactLevel.Precision:P2}/{observation.Headings.ExactLevel.Recall:P2}; expected 100%. {observation.Headings.Mismatches.FirstOrDefault()}");
        }
        RequirePerfect("table detection", observation.TableDetection);
        RequirePerfect("table cell adjacency", observation.CellAdjacency);
        RequirePerfect("cross-page table continuation", observation.ContinuationPairs);
    }

    private static void RequirePerfect(string name, PdfBinaryClassificationScore score) {
        if (score.Precision < 1D || score.Recall < 1D) {
            throw new InvalidDataException(
                $"Deterministic {name} precision/recall was {score.Precision:P2}/{score.Recall:P2}; expected 100%.");
        }
    }

    private static List<(int PageNumber, IReadOnlyList<string> Cells, IReadOnlyList<IReadOnlyList<string>> Rows)> ReadTables(
        PdfDocumentReadResult document) {
        var tables = new List<(int, IReadOnlyList<string>, IReadOnlyList<IReadOnlyList<string>>)>(document.Tables.Count);
        foreach (PdfLogicalTable table in document.Tables) {
            PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
            IReadOnlyList<IReadOnlyList<string>> rows = new[] { data.Columns }
                .Concat(data.Rows)
                .ToArray();
            tables.Add((table.PageNumber, rows.SelectMany(static row => row).ToArray(), rows));
        }
        return tables;
    }

    private static void AddAdjacencyEdges(
        Dictionary<string, int> edges,
        int pageNumber,
        int tableIndex,
        IReadOnlyList<IReadOnlyList<string>> rows) {
        for (int row = 0; row < rows.Count; row++) {
            for (int column = 0; column < rows[row].Count; column++) {
                string current = rows[row][column];
                if (column + 1 < rows[row].Count) {
                    AddEdge($"{pageNumber}:{tableIndex}:{row}:{column}:H:{current}>{rows[row][column + 1]}");
                }
                if (row + 1 < rows.Count && column < rows[row + 1].Count) {
                    AddEdge($"{pageNumber}:{tableIndex}:{row}:{column}:V:{current}>{rows[row + 1][column]}");
                }
            }
        }

        void AddEdge(string edge) {
            edges.TryGetValue(edge, out int count);
            edges[edge] = checked(count + 1);
        }
    }

    private static PdfBinaryClassificationScore Score(int truePositive, int falsePositive, int falseNegative) {
        truePositive = Math.Max(0, truePositive);
        falsePositive = Math.Max(0, falsePositive);
        falseNegative = Math.Max(0, falseNegative);
        double precision = truePositive + falsePositive == 0
            ? falseNegative == 0 ? 1D : 0D
            : (double)truePositive / (truePositive + falsePositive);
        double recall = truePositive + falseNegative == 0
            ? falsePositive == 0 ? 1D : 0D
            : (double)truePositive / (truePositive + falseNegative);
        double f1 = precision + recall == 0D ? 0D : 2D * precision * recall / (precision + recall);
        return new PdfBinaryClassificationScore(truePositive, falsePositive, falseNegative, precision, recall, f1);
    }

    private static int FindContainingIndex(IEnumerable<string> values, string marker) {
        int index = 0;
        foreach (string value in values) {
            if (value.Contains(marker, StringComparison.Ordinal)) return index;
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
                int substitution = previous[actualIndex - 1] +
                    (expected[expectedIndex - 1] == actual[actualIndex - 1] ? 0 : 1);
                current[actualIndex] = Math.Min(
                    Math.Min(previous[actualIndex] + 1, current[actualIndex - 1] + 1),
                    substitution);
            }
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }
}
