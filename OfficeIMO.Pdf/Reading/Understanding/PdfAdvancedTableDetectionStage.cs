namespace OfficeIMO.Pdf;

/// <summary>Canonical native and OCR table-evidence stage.</summary>
internal sealed class PdfAdvancedTableDetectionStage : IPdfTableDetectionStage {
    public IReadOnlyList<PdfUnderstandingTableCandidate> DetectTables(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> lines) {
        if (lines.Count == 0) return Array.Empty<PdfUnderstandingTableCandidate>();

        var result = new List<PdfUnderstandingTableCandidate>();
        List<TextLayoutEngine.TextLine> nativeLines = BuildNativeTableLines(context, lines);
        if (nativeLines.Count > 0) AddNativeCandidates(context, lines, nativeLines, result);
        AddOcrCandidates(context, lines, result);
        return result.Count == 0 ? Array.Empty<PdfUnderstandingTableCandidate>() : result.AsReadOnly();
    }

    private static void AddNativeCandidates(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> understandingLines,
        List<TextLayoutEngine.TextLine> nativeLines,
        List<PdfUnderstandingTableCandidate> result) {
        List<List<TextLayoutEngine.TextLine>> bands = TextLayoutEngine.BandLines(
            nativeLines,
            context.LayoutOptions.ToEngineOptions(),
            context.ConsumeWork,
            context.ThrowIfCancellationRequested);
        List<StructuredTable> tables = TableDetector.DetectTablesFromBands(
            bands,
            context.Height,
            context.ConsumeWork,
            context.ThrowIfCancellationRequested);
        result.Capacity = tables.Count;
        foreach (StructuredTable table in tables) {
            context.ConsumeWork();
            ContentStructureExtractor.NormalizeDetectedTable(table);
            if (table.Columns.Count < 2 || table.SourceRuns.Count == 0) continue;

            var ownedRuns = new HashSet<PdfTextSpan>();
            for (int runIndex = 0; runIndex < table.SourceRuns.Count; runIndex++) {
                context.ConsumeWork();
                ownedRuns.Add(table.SourceRuns[runIndex]);
            }
            var matchedLines = new List<PdfUnderstandingLine>();
            for (int lineIndex = 0; lineIndex < understandingLines.Count; lineIndex++) {
                context.ConsumeWork();
                PdfUnderstandingLine line = understandingLines[lineIndex];
                if (line.SourceKind != PdfLogicalContentSourceKind.Native) continue;
                var ownedWords = new List<PdfUnderstandingWord>();
                for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                    PdfUnderstandingWord word = line.Words[wordIndex];
                    IReadOnlyList<PdfTextSpan> sourceRuns = word.SourceRuns;
                    bool owned = false;
                    for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                        context.ConsumeWork();
                        if (!ownedRuns.Contains(sourceRuns[runIndex])) continue;
                        owned = true;
                        break;
                    }
                    if (owned) ownedWords.Add(word);
                }
                if (ownedWords.Count > 0) {
                    matchedLines.Add(PdfAdvancedUnderstandingStages.CreateLineSubset(
                        context,
                        line,
                        ownedWords));
                }
            }
            PdfUnderstandingLine[] sourceLines = PdfAdvancedUnderstandingStages.CopyAndSort(
                context,
                matchedLines,
                static (first, second) => {
                    int baseline = second.BaselineY.CompareTo(first.BaselineY);
                    return baseline != 0 ? baseline : first.XStart.CompareTo(second.XStart);
                });
            if (sourceLines.Length == 0 ||
                (sourceLines.Length < 2 && !string.Equals(table.Kind, "leaders", StringComparison.Ordinal))) continue;

            double confidence = PdfInference.Clamp(sourceLines.Average(static line => line.Confidence));
            var evidence = new List<PdfInferenceEvidence> {
                new PdfInferenceEvidence(
                    "table.aligned-geometry",
                    "Repeated column geometry and row alignment form a bounded table candidate.",
                    0.9D)
            };
            if (TableDetector.HasDistinctEmphasizedHeader(sourceLines)) {
                evidence.Add(new PdfInferenceEvidence(
                    "table.header-emphasis",
                    "The first row has a distinct emphasized font profile relative to the body rows.",
                    0.9D));
            }
            result.Add(PdfUnderstandingTableCandidate.FromStructured(
                table,
                sourceLines,
                confidence,
                evidence,
                context.ConsumeWork,
                context.ThrowIfCancellationRequested));
        }
    }

    private static void AddOcrCandidates(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> lines,
        List<PdfUnderstandingTableCandidate> result) {
        IReadOnlyList<PdfUnderstandingTableCandidate> ocrCandidates =
            PdfOcrTableEvidenceDetector.Detect(context, lines, context.MaxTableCandidatesPerPage);
        if (ocrCandidates.Count == 0) return;

        IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
            PdfUnderstandingTableCandidateReconciler.Reconcile(context.Page, result, ocrCandidates);
        if (reconciled.Count > context.MaxTableCandidatesPerPage) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.UnderstandingArtifacts,
                    context.MaxTableCandidatesPerPage,
                    reconciled.Count);
        }
        result.Clear();
        result.AddRange(reconciled);
    }

    private static List<TextLayoutEngine.TextLine> BuildNativeTableLines(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> lines) {
        bool hasNativeLines = false;
        bool hasNonNativeLines = false;
        var sourceRuns = new List<PdfTextSpan>();
        var seen = new HashSet<PdfTextSpan>();
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            context.ConsumeWork();
            PdfUnderstandingLine line = lines[lineIndex];
            if (line.SourceKind != PdfLogicalContentSourceKind.Native) {
                hasNonNativeLines = true;
                continue;
            }
            hasNativeLines = true;
            if (string.IsNullOrWhiteSpace(line.Text) || !IsHorizontalBaseline(line.RotationDegrees)) continue;

            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> wordRuns = line.Words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < wordRuns.Count; runIndex++) {
                    context.ConsumeWork();
                    PdfTextSpan sourceRun = wordRuns[runIndex];
                    if (seen.Add(sourceRun)) sourceRuns.Add(sourceRun);
                }
            }
        }

        if (!hasNativeLines) return new List<TextLayoutEngine.TextLine>();

        // A native decode also contains positioned whitespace runs. Word grouping deliberately
        // omits whitespace-only words, but the runs still carry gap geometry required by the
        // table detector. When the complete input is native, retain that evidence verbatim.
        if (!hasNonNativeLines) {
            sourceRuns.Clear();
            sourceRuns.Capacity = Math.Max(sourceRuns.Capacity, context.DecodedRuns.Count);
            for (int runIndex = 0; runIndex < context.DecodedRuns.Count; runIndex++) {
                context.ConsumeWork();
                sourceRuns.Add(context.DecodedRuns[runIndex]);
            }
        }

        // Table inference needs the original positioned runs. Reconstructing table rows from
        // already-grouped understanding lines changes the baseline and gap evidence and can
        // collapse rows emitted by Office producers. Reuse the shared layout primitive over
        // the native runs owned by the accepted lines instead.
        return TextLayoutEngine.BuildLines(
            sourceRuns,
            context.LayoutOptions.ToEngineOptions(),
            context.ConsumeWork,
            context.ThrowIfCancellationRequested);
    }

    private static bool IsHorizontalBaseline(double rotationDegrees) {
        double normalized = Math.Abs(PdfAdvancedUnderstandingStages.NormalizeAngle(rotationDegrees));
        return normalized <= 5D || normalized >= 175D;
    }
}
