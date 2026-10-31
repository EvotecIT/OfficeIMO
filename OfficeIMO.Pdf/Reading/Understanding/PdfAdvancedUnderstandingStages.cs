using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Built-in dependency-free understanding stages for rotated baselines, spatial regions, multi-column reading order, and richer business-document semantics.
/// </summary>
internal static partial class PdfAdvancedUnderstandingStages {
    /// <summary>Parser-backed positioned text decoding.</summary>
    public static IPdfGlyphDecodingStage GlyphDecoding { get; } = new AdvancedGlyphDecodingStage();
    /// <summary>Rotation-aware word grouping.</summary>
    public static IPdfWordGroupingStage WordGrouping { get; } = new AdvancedWordGroupingStage();
    /// <summary>Arbitrary-baseline line grouping.</summary>
    public static IPdfLineGroupingStage LineGrouping { get; } = new AdvancedLineGroupingStage();
    /// <summary>Language-neutral table candidate detection.</summary>
    public static IPdfTableDetectionStage TableDetection { get; } = new PdfAdvancedTableDetectionStage();
    /// <summary>Spatial connected-region segmentation.</summary>
    public static IPdfPageSegmentationStage PageSegmentation { get; } = new AdvancedPageSegmentationStage();
    /// <summary>Spanning-band and multi-column reading order.</summary>
    public static IPdfReadingOrderStage ReadingOrder { get; } = new PdfRecursiveXyCutReadingOrderStage();
    /// <summary>Business-document semantic classification.</summary>
    public static IPdfSemanticClassificationStage SemanticClassification { get; } = new AdvancedSemanticClassificationStage();
    /// <summary>Image-region detection and evidence-based caption association.</summary>
    public static IPdfImageRegionDetectionStage ImageRegionDetection { get; } = new PdfImageRegionDetectionStage();

    internal static T[] CopyAndSort<T>(
        PdfUnderstandingPageContext context,
        IReadOnlyList<T> values,
        Comparison<T> comparison) {
        var source = new T[values.Count];
        for (int index = 0; index < values.Count; index++) {
            context.ConsumeWork();
            source[index] = values[index];
        }
        if (source.Length < 2) return source;

        var target = new T[source.Length];
        for (int width = 1; width < source.Length;) {
            for (int left = 0; left < source.Length; left += width * 2) {
                int middle = Math.Min(left + width, source.Length);
                int right = Math.Min(left + (width * 2), source.Length);
                int first = left;
                int second = middle;
                int output = left;
                while (first < middle && second < right) {
                    context.ConsumeWork();
                    target[output++] = comparison(source[first], source[second]) <= 0
                        ? source[first++]
                        : source[second++];
                }
                while (first < middle) target[output++] = source[first++];
                while (second < right) target[output++] = source[second++];
            }
            T[] swap = source;
            source = target;
            target = swap;
            if (width > source.Length / 2) break;
            width *= 2;
        }
        return source;
    }

    internal static PdfUnderstandingLine CreateLineSubset(
        PdfUnderstandingPageContext context,
        PdfUnderstandingLine source,
        IReadOnlyList<PdfUnderstandingWord> words) =>
        CreateLineSubset(source, words, context.LayoutOptions.ReadingDirection, context.ConsumeWork, context.ThrowIfCancellationRequested);

    // Logical projection also needs subsets after document-level tagged-table enrichment.
    // Reuse the original page budget and spacing rules without creating a second pipeline context.
    internal static PdfUnderstandingLine CreateLineSubset(
        PdfUnderstandingLine source,
        IReadOnlyList<PdfUnderstandingWord> words,
        PdfReadingDirection readingDirection,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        if (words.Count == 0) throw new ArgumentException("A line subset requires at least one word.", nameof(words));
        bool unchanged = words.Count == source.Words.Count;
        for (int index = 0; unchanged && index < words.Count; index++) {
            consumeWork?.Invoke(1);
            unchanged = ReferenceEquals(words[index], source.Words[index]);
        }
        if (unchanged) return source;

        PdfUnderstandingWord[] snapshot = words.ToArray();
        PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
            readingDirection,
            snapshot.OrderBy(static word => word.SourceSequence).Select(static word => word.Text));
        int? sourceSequence = snapshot.Any(static word => word.SourceSequence.HasValue)
            ? snapshot.Where(static word => word.SourceSequence.HasValue).Min(static word => word.SourceSequence!.Value)
            : source.SourceSequence;
        return new PdfUnderstandingLine(
            Array.AsReadOnly(snapshot),
            ComposeLineText(snapshot, source.RotationDegrees, direction, cancellationCheck),
            source.Confidence,
            source.Evidence,
            source.SourceKind,
            sourceSequence,
            source.BlockId,
            source.ParagraphId,
            source.LineId);
    }

    private static string ComposeLineText(
        PdfUnderstandingPageContext context,
        PdfUnderstandingWord[] words,
        double angle,
        PdfReadingDirection direction) =>
        ComposeLineText(words, angle, direction, context.ThrowIfCancellationRequested);

    private static string ComposeLineText(
        PdfUnderstandingWord[] words,
        double angle,
        PdfReadingDirection direction,
        Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        if (words.Length == 0) return string.Empty;
        var text = new System.Text.StringBuilder(words.Sum(static word => word.Text.Length) + words.Length);
        text.Append(words[0].Text);
        for (int wordIndex = 1; wordIndex < words.Length; wordIndex++) {
            PdfUnderstandingWord previous = words[wordIndex - 1];
            PdfUnderstandingWord current = words[wordIndex];
            if (NeedsSyntheticSpace(previous, current, angle, direction)) text.Append(' ');
            text.Append(current.Text);
        }
        return text.ToString();
    }

    private static bool NeedsSyntheticSpace(
        PdfUnderstandingWord previous,
        PdfUnderstandingWord current,
        double angle,
        PdfReadingDirection direction) {
        if (SharesSourceRun(previous.SourceRuns, current.SourceRuns)) return true;
        if (HasExplicitBoundarySpace(previous.SourceRuns, current.SourceRuns, direction)) return true;
        double radians = angle * Math.PI / 180D;
        double gap = Math.Max(
            GetProjectedAlongStart(previous, radians) - GetProjectedAlongEnd(current, radians),
            GetProjectedAlongStart(current, radians) - GetProjectedAlongEnd(previous, radians));
        double threshold = Math.Max(1D, Math.Min(previous.FontSize, current.FontSize) * 0.18D);
        return gap > threshold;
    }

    private static bool SharesSourceRun(
        IReadOnlyList<PdfTextSpan> left,
        IReadOnlyList<PdfTextSpan> right) {
        for (int leftIndex = 0; leftIndex < left.Count; leftIndex++) {
            for (int rightIndex = 0; rightIndex < right.Count; rightIndex++) {
                if (ReferenceEquals(left[leftIndex], right[rightIndex])) return true;
            }
        }
        return false;
    }

    private static bool HasExplicitBoundarySpace(
        IReadOnlyList<PdfTextSpan> left,
        IReadOnlyList<PdfTextSpan> right,
        PdfReadingDirection direction) {
        if (left.Count == 0 || right.Count == 0) return false;
        PdfTextSpan previous = left[left.Count - 1];
        PdfTextSpan current = right[0];
        return direction == PdfReadingDirection.RightToLeft
            ? previous.LogicalLeadingSpace ||
              current.LogicalTrailingSpace ||
              (previous.Text.Length > 0 && char.IsWhiteSpace(previous.Text[0])) ||
              (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[current.Text.Length - 1]))
            : previous.LogicalTrailingSpace ||
              current.LogicalLeadingSpace ||
              (previous.Text.Length > 0 && char.IsWhiteSpace(previous.Text[previous.Text.Length - 1])) ||
              (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[0]));
    }

    private static double ProjectAlong(PdfUnderstandingWord word, double radians) =>
        (Math.Cos(radians) * WordAnchorX(word)) + (Math.Sin(radians) * word.BaselineY);

    private static double GetProjectedAlongStart(PdfUnderstandingWord word, double radians) {
        if (TryGetProjectedAdvance(word, out _)) {
            double startX = Math.Cos(radians) >= 0D ? word.XStart : word.XEnd;
            return (Math.Cos(radians) * startX) + (Math.Sin(radians) * word.BaselineY);
        }
        double projectedHalfExtent = GetFallbackProjectedHalfExtent(word, radians);
        return ProjectAlong(word, radians) - projectedHalfExtent;
    }

    private static double GetProjectedAlongEnd(PdfUnderstandingWord word, double radians) {
        if (TryGetProjectedAdvance(word, out double advance)) return GetProjectedAlongStart(word, radians) + advance;
        double projectedHalfExtent = GetFallbackProjectedHalfExtent(word, radians);
        return ProjectAlong(word, radians) + projectedHalfExtent;
    }

    private static bool TryGetProjectedAdvance(PdfUnderstandingWord word, out double advance) {
        advance = 0D;
        if (word.Advance is double explicitAdvance && explicitAdvance > 0.001D) {
            advance = explicitAdvance;
            return true;
        }
        if (word.SourceRuns.Count != 1 || word.Text.Length == 0) return false;
        PdfTextSpan source = word.SourceRuns[0];
        string sourceText = source.Text ?? string.Empty;
        if (sourceText.Length == 0) return false;
        int sourceScalars = PdfUnicodeScalarAnalysis.CountScalars(sourceText);
        int wordScalars = PdfUnicodeScalarAnalysis.CountScalars(word.Text);
        double perScalar = source.Advance > 0D ? source.Advance / sourceScalars : word.FontSize * 0.55D;
        advance = perScalar * wordScalars;
        return advance > 0.001D;
    }

    private static double GetFallbackProjectedHalfExtent(PdfUnderstandingWord word, double radians) {
        double horizontalProjection = Math.Abs(Math.Cos(radians)) * Math.Max(0D, word.XEnd - word.XStart) / 2D;
        if (horizontalProjection > 0.001D) return horizontalProjection;
        return Math.Max(
            word.FontSize * 0.25D,
            word.FontSize * PdfUnicodeScalarAnalysis.CountScalars(word.Text) * 0.2D);
    }

    private sealed class AdvancedGlyphDecodingStage : IPdfGlyphDecodingStage {
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            context.ThrowIfCancellationRequested();
            IReadOnlyList<PdfTextSpan> source = context.Page.GetTextSpans(context.CancellationToken);
            var runs = new PdfTextSpan[source.Count];
            for (int index = 0; index < runs.Length; index++) {
                context.ConsumeWork();
                runs[index] = source[index].WithPageFontSize();
            }
            return Array.AsReadOnly(runs);
        }
    }

    private sealed class AdvancedWordGroupingStage : IPdfWordGroupingStage {
        public IReadOnlyList<PdfUnderstandingWord> GroupWords(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
            var result = new List<PdfUnderstandingWord>();
            for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                context.ConsumeWork();
                PdfTextSpan run = runs[runIndex];
                string text = run.Text ?? string.Empty;
                if (text.Length > 0) context.ConsumeWork(text.Length);
                double radians = run.RotationDegrees * Math.PI / 180D;
                double alongX = Math.Cos(radians);
                double alongY = Math.Sin(radians);
                bool hasResolvedCharacterAdvances = PdfTextAdvanceProjection.TryGetResolvedBoundaries(run, out double[] characterBoundaries);
                if (!hasResolvedCharacterAdvances) {
                    characterBoundaries = BuildUniformScalarBoundaries(text, run.Advance, run.FontSize);
                }
                int cursor = 0;
                while (cursor < text.Length) {
                    while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) {
                        if ((cursor & 1023) == 0) context.ThrowIfCancellationRequested();
                        cursor++;
                    }
                    int start = cursor;
                    while (cursor < text.Length && !char.IsWhiteSpace(text[cursor])) {
                        if ((cursor & 1023) == 0) context.ThrowIfCancellationRequested();
                        cursor++;
                    }
                    if (cursor == start) continue;
                    context.ConsumeWork();
                    if (result.Count >= context.MaxWordsPerPage) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, context.MaxWordsPerPage, result.Count + 1L);
                    }
                    double startDistance = characterBoundaries![start];
                    double endDistance = characterBoundaries[cursor];
                    double startX = run.X + alongX * startDistance;
                    double startY = run.Y + alongY * startDistance;
                    double endX = run.X + alongX * endDistance;
                    double confidence = Math.Abs(run.RotationDegrees) <= 0.5D ? 0.96D : 0.9D;
                    int sourceSequence = result.Count;
                    string wordText = PdfTextDirectionAnalysis.RestoreLogicalOrderFromGlyphPaintSequence(
                        text.Substring(start, cursor - start),
                        run.GlyphSequenceProgressesLeftToRight);
                    if (start == 0 && TryJoinKernedFragment(context, result, runs, runIndex, run, wordText, startX, startY, endX, endDistance - startDistance)) continue;
                    result.Add(new PdfUnderstandingWord(
                        wordText,
                        Math.Min(startX, endX),
                        Math.Max(startX, endX),
                        startY,
                        run.FontSize,
                        NormalizeAngle(run.RotationDegrees),
                        new[] { run },
                        confidence,
                        new[] { new PdfInferenceEvidence(
                            hasResolvedCharacterAdvances ? "word-character-advance-projection" : "word-baseline-projection",
                            hasResolvedCharacterAdvances
                                ? "Word geometry was projected from decoded per-character advances along a " + run.RotationDegrees.ToString("0.###", CultureInfo.InvariantCulture) + " degree baseline."
                                : "Word geometry was projected uniformly along a " + run.RotationDegrees.ToString("0.###", CultureInfo.InvariantCulture) + " degree baseline because per-character advances were unavailable.",
                            hasResolvedCharacterAdvances ? 0.95D : (Math.Abs(run.RotationDegrees) <= 0.5D ? 0.8D : 0.6D)) },
                        Math.Max(0D, endDistance - startDistance),
                        visualBounds: run.HasActualText && run.TextRenderingMode == 3
                            ? GetSelectionBounds(context, startX, startY, endDistance - startDistance, run.FontSize, run.RotationDegrees)
                            : null,
                        sourceSequence: sourceSequence) { IsSelectionBox = run.HasActualText && run.TextRenderingMode == 3 });
                }
            }
            return result.Count == 0 ? Array.Empty<PdfUnderstandingWord>() : result.AsReadOnly();
        }

        // Kerning and glyph positioning split one TJ array into several spans. A fragment that
        // continues the previous span's last word on the same baseline, with no whitespace and a
        // gap well below a word space, belongs to that word.
        private static bool TryJoinKernedFragment(PdfUnderstandingPageContext context, List<PdfUnderstandingWord> result, IReadOnlyList<PdfTextSpan> runs, int runIndex,
            PdfTextSpan run, string wordText, double startX, double startY, double endX, double advance) {
            if (result.Count == 0 || runIndex == 0) return false;
            PdfTextSpan previousRun = runs[runIndex - 1];
            PdfUnderstandingWord previous = result[result.Count - 1];
            string previousText = previousRun.Text ?? string.Empty;
            if (previous.SourceRuns.Count == 0 || !ReferenceEquals(previous.SourceRuns[previous.SourceRuns.Count - 1], previousRun) ||
                previousText.Length == 0 || char.IsWhiteSpace(previousText[previousText.Length - 1]) ||
                !run.GlyphSequenceProgressesLeftToRight || !previousRun.GlyphSequenceProgressesLeftToRight ||
                Math.Abs(run.RotationDegrees) > 0.5D || Math.Abs(previousRun.RotationDegrees) > 0.5D ||
                previous.IsSelectionBox || (run.HasActualText && run.TextRenderingMode == 3) ||
                Math.Abs(previous.BaselineY - startY) > Math.Max(0.01D, run.FontSize * 0.02D) ||
                Math.Abs(previous.FontSize - run.FontSize) > Math.Max(0.01D, run.FontSize * 0.02D)) return false;
            double gap = Math.Min(startX, endX) - previous.XEnd;
            double em = Math.Max(run.FontSize, 0.01D);
            if (gap < -0.3D * em || gap > 0.1D * em) return false;
            // Joining copies the growing text and source list, then copies evidence twice.
            // Charge that work before allocating so one word cannot hide quadratic cost.
            context.ConsumeWork((long)previous.Text.Length + wordText.Length + previous.SourceRuns.Count +
                (long)previous.Evidence.Count * 2L + 3L);
            var sources = new List<PdfTextSpan>(previous.SourceRuns) { run };
            var evidence = new List<PdfInferenceEvidence>(previous.Evidence) {
                new PdfInferenceEvidence("word-kerned-fragment-join",
                    "Adjacent same-baseline glyph runs without whitespace were joined into one word.", 0.9D)
            };
            result[result.Count - 1] = new PdfUnderstandingWord(
                previous.Text + wordText,
                Math.Min(previous.XStart, Math.Min(startX, endX)),
                Math.Max(previous.XEnd, Math.Max(startX, endX)),
                previous.BaselineY,
                previous.FontSize,
                previous.RotationDegrees,
                sources,
                Math.Min(previous.Confidence, 0.95D),
                evidence,
                (previous.Advance ?? 0D) + gap + advance,
                sourceSequence: previous.SourceSequence);
            return true;
        }

        private static double[] BuildUniformScalarBoundaries(string text, double advance, double fontSize) {
            var boundaries = new double[text.Length + 1];
            int scalarCount = PdfUnicodeScalarAnalysis.CountScalars(text);
            if (scalarCount == 0) return boundaries;
            double scalarAdvance = advance > 0D ? advance / scalarCount : fontSize * 0.55D;
            double distance = 0D;
            for (int index = 0; index < text.Length;) {
                boundaries[index] = distance;
                int scalarLength = char.IsSurrogatePair(text, index) ? 2 : 1;
                if (scalarLength == 2) boundaries[index + 1] = distance;
                distance += scalarAdvance;
                index += scalarLength;
                boundaries[index] = distance;
            }
            return boundaries;
        }
    }

    private sealed class AdvancedPageSegmentationStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            lines = BuildSegmentationLines(context, lines);
            var remaining = new HashSet<int>(Enumerable.Range(0, lines.Count));
            var regions = new List<PdfUnderstandingRegion>();
            AddCanonicalTableRegions(context, lines, remaining, regions);
            double pageMedianFontSize = GetLowerMedianFontSize(remaining.Select(index => lines[index]));
            int[] parent = Enumerable.Range(0, lines.Count).ToArray();
            double maximumVerticalGap = lines.Count == 0 ? 0D : lines.Max(static line => line.FontSize) * 2.2D;
            int[] candidates = CopyAndSort(
                context,
                remaining.ToArray(),
                (left, right) => {
                    int baseline = lines[right].BaselineY.CompareTo(lines[left].BaselineY);
                    if (baseline != 0) return baseline;
                    int x = lines[left].XStart.CompareTo(lines[right].XStart);
                    return x != 0 ? x : left.CompareTo(right);
                });
            for (int leftPosition = 0; leftPosition < candidates.Length; leftPosition++) {
                int leftIndex = candidates[leftPosition];
                for (int rightPosition = leftPosition + 1; rightPosition < candidates.Length; rightPosition++) {
                    int rightIndex = candidates[rightPosition];
                    double verticalGap = lines[leftIndex].BaselineY - lines[rightIndex].BaselineY;
                    if (verticalGap > maximumVerticalGap) break;
                    context.ConsumeWork();
                    if (AreSpatialNeighbors(lines[leftIndex], lines[rightIndex])) Union(parent, leftIndex, rightIndex);
                }
            }
            foreach (IGrouping<int, int> component in candidates.GroupBy(index => Find(parent, index))) {
                PdfUnderstandingLine[] ordered = CopyAndSort(
                    context,
                    component.Select(index => lines[index]).ToArray(),
                    static (left, right) => {
                        if (left.SourceKind == PdfLogicalContentSourceKind.Ocr &&
                            right.SourceKind == PdfLogicalContentSourceKind.Ocr &&
                            left.SourceSequence.HasValue && right.SourceSequence.HasValue) {
                            int source = left.SourceSequence.Value.CompareTo(right.SourceSequence.Value);
                            if (source != 0) return source;
                        }
                        int baseline = right.BaselineY.CompareTo(left.BaselineY);
                        return baseline != 0 ? baseline : left.XStart.CompareTo(right.XStart);
                    });
                foreach (PdfUnderstandingLine[] column in SplitSideBySideColumns(context, ordered, depth: 0)) {
                    AddStructuralSubregions(context, column, pageMedianFontSize, regions);
                }
            }
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
        }

        /// <summary>
        /// Splits a connected component at a vertical gutter when it holds side-by-side text columns. Lines from
        /// neighbouring columns share baselines and can connect through a spanning title, so ordering the component by
        /// baseline alone would interleave the columns. Lines that cross the gutter stay as their own bands.
        /// </summary>
        private static IEnumerable<PdfUnderstandingLine[]> SplitSideBySideColumns(
            PdfUnderstandingPageContext context,
            PdfUnderstandingLine[] component,
            int depth) {
            const int MinimumColumnLines = 3;
            const double GutterTolerance = 1.5D;
            if (depth >= 8 || component.Length < MinimumColumnLines * 2 ||
                component.Any(static line => line.SourceKind == PdfLogicalContentSourceKind.Ocr && line.SourceSequence.HasValue)) {
                yield return component;
                yield break;
            }
            double left = component.Min(static line => line.XStart);
            double right = component.Max(static line => line.XEnd);
            double width = right - left;
            double[] fontSizes = component.Select(static line => Math.Max(1D, line.FontSize)).OrderBy(static size => size).ToArray();
            double medianFont = fontSizes[fontSizes.Length / 2];
            double minimumGutter = Math.Max(6D, medianFont * 0.8D);
            if (width <= minimumGutter * 4D) {
                yield return component;
                yield break;
            }
            // Find the widest vertical strip crossed by only a few lines. Titles, bylines and the odd merged line may
            // cross a real gutter; they become spanning bands below.
            int bins = (int)Math.Ceiling(width) + 1;
            if (bins > 20_000) {
                yield return component;
                yield break;
            }
            context.ConsumeWork(component.Length + bins);
            var coverage = new int[bins];
            for (int index = 0; index < component.Length; index++) {
                int start = Math.Max(0, (int)Math.Floor(component[index].XStart - left));
                int end = Math.Min(bins - 1, (int)Math.Ceiling(component[index].XEnd - left));
                for (int bin = start; bin <= end; bin++) coverage[bin]++;
            }
            int allowance = Math.Max(1, component.Length / 10);
            int bestStart = -1, bestLength = 0, runStart = -1;
            for (int bin = 0; bin <= bins; bin++) {
                bool open = bin < bins && coverage[bin] <= allowance;
                if (open) {
                    if (runStart < 0) runStart = bin;
                    continue;
                }
                if (runStart >= 0 && runStart > 0 && bin < bins && bin - runStart > bestLength) {
                    bestStart = runStart;
                    bestLength = bin - runStart;
                }
                runStart = -1;
            }
            if (bestStart < 0 || bestLength < minimumGutter) {
                yield return component;
                yield break;
            }
            double gutterStart = left + bestStart;
            double gutterEnd = left + bestStart + bestLength;
            var leftLines = new List<PdfUnderstandingLine>();
            var rightLines = new List<PdfUnderstandingLine>();
            for (int index = 0; index < component.Length; index++) {
                PdfUnderstandingLine line = component[index];
                if (line.XEnd <= gutterStart + GutterTolerance) leftLines.Add(line);
                else if (line.XStart >= gutterEnd - GutterTolerance) rightLines.Add(line);
            }
            if (!IsTextColumn(leftLines, width, medianFont) || !IsTextColumn(rightLines, width, medianFont) ||
                !ShareVerticalExtent(leftLines, rightLines, medianFont)) {
                yield return component;
                yield break;
            }
            // Emit bands top to bottom: a spanning line closes the current column band.
            var currentLeft = new List<PdfUnderstandingLine>();
            var currentRight = new List<PdfUnderstandingLine>();
            for (int index = 0; index < component.Length; index++) {
                PdfUnderstandingLine line = component[index];
                if (line.XEnd <= gutterStart + GutterTolerance) { currentLeft.Add(line); continue; }
                if (line.XStart >= gutterEnd - GutterTolerance) { currentRight.Add(line); continue; }
                // A baseline shared by both columns can join their words into one line; split it at the gutter.
                PdfUnderstandingWord[] leftWords = line.Words.Where(word => word.XEnd <= gutterStart + GutterTolerance).ToArray();
                PdfUnderstandingWord[] rightWords = line.Words.Where(word => word.XStart >= gutterEnd - GutterTolerance).ToArray();
                if (leftWords.Length > 0 && rightWords.Length > 0 && leftWords.Length + rightWords.Length == line.Words.Count) {
                    currentLeft.Add(CreateLineSubset(context, line, leftWords));
                    currentRight.Add(CreateLineSubset(context, line, rightWords));
                    continue;
                }
                foreach (PdfUnderstandingLine[] band in FlushColumns()) yield return band;
                yield return new[] { line };
            }
            foreach (PdfUnderstandingLine[] band in FlushColumns()) yield return band;

            IEnumerable<PdfUnderstandingLine[]> FlushColumns() {
                var bands = new List<PdfUnderstandingLine[]>();
                if (currentLeft.Count > 0) bands.AddRange(SplitSideBySideColumns(context, currentLeft.ToArray(), depth + 1));
                if (currentRight.Count > 0) bands.AddRange(SplitSideBySideColumns(context, currentRight.ToArray(), depth + 1));
                currentLeft = new List<PdfUnderstandingLine>();
                currentRight = new List<PdfUnderstandingLine>();
                return bands;
            }
        }

        private static bool IsTextColumn(List<PdfUnderstandingLine> lines, double componentWidth, double medianFont) {
            if (lines.Count < 3) return false;
            double[] widths = lines.Select(static line => line.XEnd - line.XStart).OrderBy(static value => value).ToArray();
            double medianWidth = widths[widths.Length / 2];
            // Short label/value pairs are not prose columns; require lines several words wide.
            return medianWidth >= Math.Max(medianFont * 8D, componentWidth * 0.2D);
        }

        private static bool ShareVerticalExtent(
            List<PdfUnderstandingLine> left,
            List<PdfUnderstandingLine> right,
            double medianFont) {
            double leftTop = left.Max(static line => line.BaselineY);
            double leftBottom = left.Min(static line => line.BaselineY);
            double rightTop = right.Max(static line => line.BaselineY);
            double rightBottom = right.Min(static line => line.BaselineY);
            double overlap = Math.Min(leftTop, rightTop) - Math.Max(leftBottom, rightBottom);
            return overlap >= medianFont * 2D;
        }

        private static double GetLowerMedianFontSize(IEnumerable<PdfUnderstandingLine> lines) {
            double[] sizes = lines.Select(static line => line.FontSize).OrderBy(static size => size).ToArray();
            return sizes.Length == 0 ? 0D : sizes[(sizes.Length - 1) / 2];
        }

        private static void AddStructuralSubregions(
            PdfUnderstandingPageContext context,
            PdfUnderstandingLine[] component,
            double pageMedianFontSize,
            List<PdfUnderstandingRegion> regions) {
            var current = new List<PdfUnderstandingLine>();
            for (int lineIndex = 0; lineIndex < component.Length; lineIndex++) {
                context.ConsumeWork();
                PdfUnderstandingLine line = component[lineIndex];
                bool startsHeading = IsTypographicallyProminent(line, pageMedianFontSize);
                bool startsListItem = ContentStructureExtractor.IsListItemText(line.Text.Trim());
                if (current.Count > 0) {
                    PdfUnderstandingLine first = current[0];
                    PdfUnderstandingLine previous = current[current.Count - 1];
                    bool currentIsHeading = IsTypographicallyProminent(first, pageMedianFontSize);
                    bool currentIsListItem = ContentStructureExtractor.IsListItemText(first.Text.Trim());
                    bool continuesListItem = currentIsListItem && IsWrappedListContinuation(first, previous, line);
                    bool continuesHeading = currentIsHeading && startsHeading && !startsListItem &&
                        IsWrappedHeadingContinuation(current, line);
                    if (continuesHeading) {
                        current.Add(line);
                        continue;
                    }
                    if (startsHeading || startsListItem || currentIsHeading ||
                        HasMaterialFontTransition(previous, line) ||
                        (!currentIsListItem && StartsParagraphBoundary(current, line)) ||
                        (currentIsListItem && !continuesListItem)) {
                        AddRegion(current, regions);
                        current = new List<PdfUnderstandingLine>();
                    }
                }
                current.Add(line);
            }
            AddRegion(current, regions);
        }

        /// <summary>
        /// A display title that wraps (beside a logo, or in a narrow column) continues on the next line in the same
        /// size and margin at single line spacing; the line before does not end a sentence and is not shorter.
        /// </summary>
        private static bool IsWrappedHeadingContinuation(List<PdfUnderstandingLine> current, PdfUnderstandingLine line) {
            if (current.Count >= 3) return false;
            PdfUnderstandingLine previous = current[current.Count - 1];
            if (Math.Abs(previous.FontSize - line.FontSize) > 0.5D) return false;
            if (Math.Abs(previous.XStart - line.XStart) > Math.Max(2D, line.FontSize * 0.25D)) return false;
            double gap = previous.BaselineY - line.BaselineY;
            if (gap <= 0D || gap > line.FontSize * 1.45D) return false;
            // Wrapping fills the first line; a short line over a longer one is a separate heading above it.
            if (previous.XEnd - previous.XStart < (line.XEnd - line.XStart) * 0.8D) return false;
            string text = previous.Text.Trim();
            return text.Length > 0 && !ContentStructureExtractor.EndsWithSentenceTerminal(text);
        }

        private static bool IsTypographicallyProminent(PdfUnderstandingLine line, double pageMedianFontSize) =>
            pageMedianFontSize > 0D && line.FontSize >= pageMedianFontSize * 1.2D;

        private static bool HasMaterialFontTransition(PdfUnderstandingLine previous, PdfUnderstandingLine current) {
            if (previous.FontSize <= 0D || current.FontSize <= 0D) return false;
            return Math.Max(previous.FontSize, current.FontSize) / Math.Min(previous.FontSize, current.FontSize) >= 1.2D;
        }

        private static bool StartsParagraphBoundary(
            List<PdfUnderstandingLine> current,
            PdfUnderstandingLine candidate) {
            if (HasExplicitParagraphBreak(candidate)) return true;
            if (current.Count == 0) return false;

            PdfUnderstandingLine first = current[0];
            PdfUnderstandingLine previous = current[current.Count - 1];
            if (HasProviderHierarchyBoundary(previous, candidate)) return true;
            double verticalGap = previous.BaselineY - candidate.BaselineY;
            if (verticalGap <= 0D) return false;
            double em = Math.Max(1D, Math.Max(previous.FontSize, candidate.FontSize));
            double indentation = em * 0.75D;
            if (current.Count == 1) {
                if (verticalGap >= em * 1.75D) return true;
                return verticalGap >= em * 1.15D &&
                    Math.Abs(candidate.XStart - previous.XStart) >= indentation;
            }

            double previousGap = current[current.Count - 2].BaselineY - previous.BaselineY;
            if (previousGap > 0D &&
                verticalGap >= Math.Max(em * 1.5D, previousGap * 1.45D)) return true;

            if (verticalGap < em * 1.15D) return false;
            bool startsIndentedLine = candidate.XStart - previous.XStart >= indentation &&
                Math.Abs(previous.XStart - first.XStart) <= indentation * 0.5D;
            bool returnsToParagraphMargin = Math.Abs(candidate.XStart - first.XStart) <= indentation * 0.5D &&
                Math.Abs(previous.XStart - first.XStart) >= indentation;
            return startsIndentedLine || returnsToParagraphMargin;
        }

        private static bool HasProviderHierarchyBoundary(
            PdfUnderstandingLine previous,
            PdfUnderstandingLine candidate) {
            if (previous.SourceKind != PdfLogicalContentSourceKind.Ocr ||
                candidate.SourceKind != PdfLogicalContentSourceKind.Ocr) return false;
            bool blockBoundary = (previous.BlockId is not null || candidate.BlockId is not null) &&
                !string.Equals(previous.BlockId, candidate.BlockId, StringComparison.Ordinal);
            bool paragraphBoundary = (previous.ParagraphId is not null || candidate.ParagraphId is not null) &&
                !string.Equals(previous.ParagraphId, candidate.ParagraphId, StringComparison.Ordinal);
            return blockBoundary || paragraphBoundary;
        }

        private static bool HasExplicitParagraphBreak(PdfUnderstandingLine line) {
            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> sourceRuns = line.Words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                    if (sourceRuns[runIndex].LogicalLineBreaksBefore >= 2) return true;
                }
            }
            return false;
        }

        private static bool IsWrappedListContinuation(
            PdfUnderstandingLine first,
            PdfUnderstandingLine previous,
            PdfUnderstandingLine candidate) {
            if (ContentStructureExtractor.IsListItemText(candidate.Text.Trim())) return false;
            double indentation = candidate.XStart - first.XStart;
            double minimumIndentation = Math.Max(1.5D, Math.Min(first.FontSize, candidate.FontSize) * 0.25D);
            double verticalGap = previous.BaselineY - candidate.BaselineY;
            double maximumVerticalGap = Math.Max(first.FontSize, candidate.FontSize) * 1.8D;
            double fontRatio = first.FontSize <= 0D || candidate.FontSize <= 0D
                ? 1D
                : Math.Max(first.FontSize, candidate.FontSize) / Math.Min(first.FontSize, candidate.FontSize);
            return indentation >= minimumIndentation &&
                   verticalGap >= 0D &&
                   verticalGap <= maximumVerticalGap &&
                   fontRatio <= 1.2D;
        }

        private static void AddRegion(
            List<PdfUnderstandingLine> lines,
            List<PdfUnderstandingRegion> regions) {
            if (lines.Count == 0) return;
            double confidence = PdfInference.Clamp(lines.Average(static line => line.Confidence) - Math.Min(0.2D, Math.Max(0, lines.Count - 12) * 0.01D));
            regions.Add(new PdfUnderstandingRegion(lines.AsReadOnly(), confidence, new[] {
                new PdfInferenceEvidence(
                    "region-structural-connectivity",
                    "The region is a spatially connected run bounded by typographic transitions and language-neutral list syntax.",
                    lines.Count > 1 ? 0.8D : 0.6D)
            }));
        }

        private static IReadOnlyList<PdfUnderstandingLine> BuildSegmentationLines(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingLine> lines) {
            PdfUnderstandingTableCandidate[] tables = context.TableCandidates
                .Where(static table => !string.Equals(table.DetectionKind, "leaders", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (tables.Length == 0) return lines;

            var ownedRuns = new HashSet<PdfTextSpan>();
            var tableLines = new HashSet<PdfUnderstandingLine>();
            for (int tableIndex = 0; tableIndex < tables.Length; tableIndex++) {
                IReadOnlyList<PdfUnderstandingLine> sourceLines = tables[tableIndex].SourceLines;
                for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
                    context.ConsumeWork();
                    PdfUnderstandingLine sourceLine = sourceLines[lineIndex];
                    tableLines.Add(sourceLine);
                    for (int wordIndex = 0; wordIndex < sourceLine.Words.Count; wordIndex++) {
                        IReadOnlyList<PdfTextSpan> sourceRuns = sourceLine.Words[wordIndex].SourceRuns;
                        for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                            context.ConsumeWork();
                            ownedRuns.Add(sourceRuns[runIndex]);
                        }
                    }
                }
            }

            var result = new List<PdfUnderstandingLine>(lines.Count + tableLines.Count);
            result.AddRange(tableLines);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                context.ConsumeWork();
                PdfUnderstandingLine line = lines[lineIndex];
                if (tableLines.Contains(line)) continue;
                var unownedWords = new List<PdfUnderstandingWord>(line.Words.Count);
                for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                    PdfUnderstandingWord word = line.Words[wordIndex];
                    bool owned = false;
                    for (int runIndex = 0; runIndex < word.SourceRuns.Count; runIndex++) {
                        context.ConsumeWork();
                        if (!ownedRuns.Contains(word.SourceRuns[runIndex])) continue;
                        owned = true;
                        break;
                    }
                    if (!owned) unownedWords.Add(word);
                }
                if (unownedWords.Count == line.Words.Count) {
                    result.Add(line);
                } else if (unownedWords.Count > 0) {
                    result.Add(CreateLineSubset(context, line, unownedWords));
                }
            }
            PdfUnderstandingLine[] ordered = CopyAndSort(
                context,
                result.Distinct().ToArray(),
                static (left, right) => {
                    int baseline = right.BaselineY.CompareTo(left.BaselineY);
                    return baseline != 0 ? baseline : left.XStart.CompareTo(right.XStart);
                });
            return ordered.Length == 0 ? Array.Empty<PdfUnderstandingLine>() : Array.AsReadOnly(ordered);
        }

        private static void AddCanonicalTableRegions(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingLine> lines,
            HashSet<int> remaining,
            List<PdfUnderstandingRegion> regions) {
            if (context.TableCandidates.Count == 0 || lines.Count == 0) return;

            foreach (PdfUnderstandingTableCandidate table in context.TableCandidates) {
                context.ConsumeWork();
                if (string.Equals(table.DetectionKind, "leaders", StringComparison.OrdinalIgnoreCase)) continue;
                var sourceLines = new HashSet<PdfUnderstandingLine>(table.SourceLines);
                var matchedIndexes = new List<int>();
                foreach (int index in remaining) {
                    context.ConsumeWork();
                    if (sourceLines.Contains(lines[index])) matchedIndexes.Add(index);
                }
                int[] tableLineIndexes = CopyAndSort(
                    context,
                    matchedIndexes,
                    (first, second) => {
                        PdfUnderstandingLine firstLine = lines[first];
                        PdfUnderstandingLine secondLine = lines[second];
                        if (firstLine.SourceKind == PdfLogicalContentSourceKind.Ocr &&
                            secondLine.SourceKind == PdfLogicalContentSourceKind.Ocr &&
                            firstLine.SourceSequence.HasValue && secondLine.SourceSequence.HasValue) {
                            int source = firstLine.SourceSequence.Value.CompareTo(secondLine.SourceSequence.Value);
                            if (source != 0) return source;
                        }
                        int baseline = lines[second].BaselineY.CompareTo(lines[first].BaselineY);
                        return baseline != 0 ? baseline : lines[first].XStart.CompareTo(lines[second].XStart);
                    });
                if (tableLineIndexes.Length < 2) continue;

                PdfUnderstandingLine[] tableLines = tableLineIndexes.Select(index => lines[index]).ToArray();
                foreach (int index in tableLineIndexes) remaining.Remove(index);
                double confidence = PdfInference.Clamp(tableLines.Average(static line => line.Confidence));
                regions.Add(new PdfUnderstandingRegion(tableLines, confidence, new[] {
                    new PdfInferenceEvidence(
                        "region-canonical-table",
                        "The page table-detection stage owns these aligned source lines.",
                        0.9D)
                }));
            }
        }

        private static int Find(int[] parent, int value) {
            while (parent[value] != value) {
                parent[value] = parent[parent[value]];
                value = parent[value];
            }
            return value;
        }

        private static void Union(int[] parent, int left, int right) {
            int leftRoot = Find(parent, left);
            int rightRoot = Find(parent, right);
            if (leftRoot != rightRoot) parent[rightRoot] = leftRoot;
        }

        private static bool AreSpatialNeighbors(PdfUnderstandingLine left, PdfUnderstandingLine right) {
            if (AngularDistance(left.RotationDegrees, right.RotationDegrees) > 4D) return false;
            double verticalGap = Math.Abs(left.BaselineY - right.BaselineY);
            double allowedVertical = Math.Max(left.FontSize, right.FontSize) * 2.2D;
            double horizontalGap = left.XEnd < right.XStart ? right.XStart - left.XEnd : right.XEnd < left.XStart ? left.XStart - right.XEnd : 0D;
            double allowedHorizontal = Math.Max(18D, Math.Max(left.FontSize, right.FontSize) * 2D);
            bool overlapsHorizontally = left.XStart <= right.XEnd + allowedHorizontal && right.XStart <= left.XEnd + allowedHorizontal;
            return verticalGap <= allowedVertical && overlapsHorizontally && horizontalGap <= allowedHorizontal;
        }
    }

    private sealed class AdvancedSemanticClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) {
            double median = GetBodyFontSize(orderedRegions);
            var result = new List<PdfUnderstandingSemanticElement>(orderedRegions.Count);
            foreach (PdfUnderstandingRegion region in orderedRegions) {
                context.ConsumeWork();
                (PdfUnderstandingSemanticKind kind, double confidence, string code, string message) = Classify(context, region, orderedRegions, median);
                result.Add(new PdfUnderstandingSemanticElement(region, kind, confidence, new[] { new PdfInferenceEvidence(code, message, confidence - 0.5D) }));
            }
            return result.AsReadOnly();
        }

        private static (PdfUnderstandingSemanticKind Kind, double Confidence, string Code, string Message) Classify(
            PdfUnderstandingPageContext context,
            PdfUnderstandingRegion region,
            IReadOnlyList<PdfUnderstandingRegion> pageRegions,
            double median) {
            string text = region.Text.Trim();
            double largest = region.Lines.Max(static line => line.FontSize);
            if (region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region-canonical-table", StringComparison.Ordinal))) return (PdfUnderstandingSemanticKind.Table, 0.93D, "semantic.canonical-table", "The canonical layout engine recovered the region as a validated table.");
            if (IsAdjacentTableCaption(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Caption, 0.78D, "semantic.table-caption-geometry", "A compact line is immediately adjacent to and aligned with a validated table.");
            if (PdfImageRegionDetectionStage.IsAdjacentImageCaption(context, region)) return (PdfUnderstandingSemanticKind.Caption, 0.8D, "semantic.figure-caption-geometry", "A compact region is immediately adjacent to and aligned with a positioned image.");
            if (IsNumberedHeading(region, text, largest, median)) return (PdfUnderstandingSemanticKind.Heading, 0.84D, "semantic.numbered-heading", "A single short numbered line is set materially larger than the body text.");
            if (ContentStructureExtractor.IsListItemText(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, "semantic-list-marker", "The region begins with a bullet or numbered marker.");
            // Figures such as KPI values, totals and page numbers are often set large, but text without any letters
            // cannot introduce a section.
            if (!HasLetter(text)) return (PdfUnderstandingSemanticKind.Paragraph, 0.7D, "semantic.non-textual-prominent-value", "Large or prominent text without letters is a value, not a heading.");
            if (IsLocalTypographicLead(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Heading, 0.8D, "semantic.local-font-transition", "A compact line is materially larger than the nearby aligned content it introduces.");
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, 0.82D, "semantic.large-font", "The region font is materially larger than the page median.");
            if (IsCenteredSpanningBand(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Heading, 0.76D, "semantic.centered-spanning-band", "A compact centered line forms a band above or below two separated content columns.");
            return (PdfUnderstandingSemanticKind.Paragraph, 0.72D, "semantic.body-region", "No stronger business-document semantic signal was found.");
        }

        /// <summary>
        /// The page's reference (body) font size, taken from running text. Lines of validated tables are left out,
        /// and every line is weighted by its length, so a dashboard whose 9 pt chart axes, legends and table cells
        /// outnumber its 11 pt prose still measures the prose. Lines without words (axis values, KPI figures) only
        /// count when nothing else is on the page.
        /// </summary>
        internal static double GetBodyFontSize(IReadOnlyList<PdfUnderstandingRegion> regions) {
            var running = new List<PdfUnderstandingLine>();
            var all = new List<PdfUnderstandingLine>();
            foreach (PdfUnderstandingRegion region in regions) {
                bool table = region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region-canonical-table", StringComparison.Ordinal));
                foreach (PdfUnderstandingLine line in region.Lines) {
                    all.Add(line);
                    if (!table && HasLetter(line.Text)) running.Add(line);
                }
            }
            return GetBodyFontSize(running.Count > 0 ? running : all);
        }

        /// <summary>
        /// The larger of the median line size and the length-weighted median size (the size that carries half of
        /// the text). Many short labels cannot pull the body below the prose, and one long footnote cannot pull it
        /// below the ordinary lines. Pages with fewer than five lines use the length-weighted size alone.
        /// </summary>
        internal static double GetBodyFontSize(IEnumerable<PdfUnderstandingLine> lines) {
            (double Size, int Weight)[] weighted = lines
                .Select(static line => (line.FontSize, Math.Max(1, line.Text.Trim().Length)))
                .OrderBy(static item => item.Item1)
                .ToArray();
            if (weighted.Length == 0) return 0D;
            // On a short page a few large headings outvote the body line by line (two headings, one paragraph).
            if (weighted.Length < 5) return GetLengthWeightedMedian(weighted);
            double lineMedian = weighted[(weighted.Length - 1) / 2].Size;
            return Math.Max(lineMedian, GetLengthWeightedMedian(weighted));
        }

        private static double GetLengthWeightedMedian((double Size, int Weight)[] weighted) {
            long total = weighted.Sum(static item => (long)item.Weight);
            long cumulative = 0;
            foreach ((double size, int weight) in weighted) {
                cumulative += weight;
                if (cumulative * 2 >= total) return size;
            }
            return weighted[weighted.Length - 1].Size;
        }

        /// <summary>
        /// True when the text reads as words rather than a value: "1.8 h", "94.2%" or "Q3" are figures with a unit.
        /// </summary>
        /// <summary>
        /// True for a section number such as "2. Revenue by region" set in a heading size. A list item shares the
        /// marker, but it is set in the body size or reads as a sentence.
        /// </summary>
        internal static bool IsNumberedHeading(PdfUnderstandingRegion region, string text, double largest, double median) {
            if (median <= 0D || region.Lines.Count != 1 || largest < median * 1.2D) return false;
            if (!ContentStructureExtractor.TryParseListItemText(text, out string marker, out string body, out _)) return false;
            if (marker.Length == 0 || !char.IsDigit(marker[0]) || body.Length == 0) return false;
            if (PdfUnicodeScalarAnalysis.CountScalars(body) > 90 || ContentStructureExtractor.EndsWithSentenceTerminal(body)) return false;
            return HasLetter(body);
        }

        internal static bool HasLetter(string text) {
            int letters = 0, digits = 0;
            for (int index = 0; index < text.Length; index++) {
                if (char.IsLetter(text[index])) letters++;
                else if (char.IsDigit(text[index])) digits++;
            }
            return letters >= 3 || letters > 0 && letters > digits * 2;
        }

        private static bool IsAdjacentTableCaption(
            PdfUnderstandingPageContext context,
            PdfUnderstandingRegion candidate,
            IReadOnlyList<PdfUnderstandingRegion> pageRegions) {
            string text = candidate.Text.Trim();
            if (candidate.Lines.Count != 1 || text.Length == 0 || PdfUnicodeScalarAnalysis.CountScalars(text) > 320) return false;

            PdfVisualBounds candidateBounds = GetVisualBounds(context, candidate);
            double candidateWidth = candidateBounds.Right - candidateBounds.Left;
            double maximumGap = Math.Max(6D, candidate.Lines[0].FontSize * 1.75D);
            for (int regionIndex = 0; regionIndex < pageRegions.Count; regionIndex++) {
                context.ConsumeWork();
                PdfUnderstandingRegion table = pageRegions[regionIndex];
                if (ReferenceEquals(candidate, table) ||
                    !table.Evidence.Any(static evidence => string.Equals(evidence.Code, "region-canonical-table", StringComparison.Ordinal))) continue;

                PdfVisualBounds tableBounds = GetVisualBounds(context, table);
                double tableWidth = tableBounds.Right - tableBounds.Left;
                if (candidateWidth <= 0D || tableWidth <= 0D || candidateWidth > tableWidth * 1.1D) continue;
                double[] tableFontSizes = table.Lines
                    .Select(static line => line.FontSize)
                    .OrderBy(static size => size)
                    .ToArray();
                double tableMedianFontSize = tableFontSizes[(tableFontSizes.Length - 1) / 2];
                if (candidate.Lines[0].FontSize > tableMedianFontSize * 0.95D) continue;
                double overlap = Math.Min(candidateBounds.Right, tableBounds.Right) - Math.Max(candidateBounds.Left, tableBounds.Left);
                if (overlap <= 0D || overlap / candidateWidth < 0.8D) continue;

                double aboveGap = tableBounds.Top - candidateBounds.Bottom;
                double belowGap = candidateBounds.Top - tableBounds.Bottom;
                if ((aboveGap >= -1D && aboveGap <= maximumGap) ||
                    (belowGap >= -1D && belowGap <= maximumGap)) return true;
            }
            return false;
        }

        private static bool IsLocalTypographicLead(
            PdfUnderstandingPageContext context,
            PdfUnderstandingRegion candidate,
            IReadOnlyList<PdfUnderstandingRegion> pageRegions) {
            if (candidate.Lines.Count != 1 || candidate.Text.Trim().Length == 0) return false;
            PdfVisualBounds candidateBounds = GetVisualBounds(context, candidate);
            double candidateFontSize = candidate.Lines[0].FontSize;
            double bestGap = double.MaxValue;
            double introducedFontSize = 0D;
            for (int regionIndex = 0; regionIndex < pageRegions.Count; regionIndex++) {
                context.ConsumeWork();
                PdfUnderstandingRegion region = pageRegions[regionIndex];
                if (ReferenceEquals(candidate, region) ||
                    region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region-canonical-table", StringComparison.Ordinal))) continue;
                PdfVisualBounds bounds = GetVisualBounds(context, region);
                double gap = bounds.Top - candidateBounds.Bottom;
                if (gap < -1D || gap > Math.Max(18D, candidateFontSize * 4D)) continue;
                double horizontalOverlap = Math.Min(candidateBounds.Right, bounds.Right) - Math.Max(candidateBounds.Left, bounds.Left);
                bool aligned = horizontalOverlap > 0D || Math.Abs(candidateBounds.Left - bounds.Left) <= candidateFontSize * 2D;
                if (!aligned || gap >= bestGap) continue;
                bestGap = gap;
                introducedFontSize = region.Lines.Max(static line => line.FontSize);
            }
            return introducedFontSize > 0D && candidateFontSize >= introducedFontSize * 1.2D;
        }

        private static bool IsCenteredSpanningBand(
            PdfUnderstandingPageContext context,
            PdfUnderstandingRegion candidate,
            IReadOnlyList<PdfUnderstandingRegion> pageRegions) {
            if (candidate.Lines.Count != 1 || candidate.Text.Trim().Length == 0) return false;
            PdfVisualBounds candidateBounds = GetVisualBounds(context, candidate);
            (double pageWidth, _) = context.GetVisualSize();
            double candidateWidth = candidateBounds.Right - candidateBounds.Left;
            double pageCenter = pageWidth / 2D;
            double candidateCenter = (candidateBounds.Left + candidateBounds.Right) / 2D;
            if (candidateWidth <= 0D || candidateWidth > pageWidth * 0.65D ||
                candidateBounds.Left > pageCenter || candidateBounds.Right < pageCenter ||
                Math.Abs(candidateCenter - pageCenter) > Math.Max(12D, pageWidth * 0.12D)) return false;

            bool hasLeftColumn = false;
            bool hasRightColumn = false;
            for (int regionIndex = 0; regionIndex < pageRegions.Count; regionIndex++) {
                context.ConsumeWork();
                PdfUnderstandingRegion region = pageRegions[regionIndex];
                if (ReferenceEquals(candidate, region)) continue;
                PdfVisualBounds bounds = GetVisualBounds(context, region);
                bool separatedVertically = bounds.Top >= candidateBounds.Bottom || bounds.Bottom <= candidateBounds.Top;
                if (!separatedVertically) continue;
                double center = (bounds.Left + bounds.Right) / 2D;
                if (center < pageCenter - pageWidth * 0.08D) hasLeftColumn = true;
                if (center > pageCenter + pageWidth * 0.08D) hasRightColumn = true;
                if (hasLeftColumn && hasRightColumn) return true;
            }
            return false;
        }

        private static PdfVisualBounds GetVisualBounds(PdfUnderstandingPageContext context, PdfUnderstandingRegion region) {
            PdfLogicalVisualBounds[] directBounds = region.Lines
                .Select(static line => line.VisualBounds)
                .Where(static bounds => bounds is not null)
                .Cast<PdfLogicalVisualBounds>()
                .ToArray();
            if (directBounds.Length == region.Lines.Count) {
                return new PdfVisualBounds(
                    directBounds.Min(static bounds => bounds.Left),
                    directBounds.Min(static bounds => bounds.Top),
                    directBounds.Max(static bounds => bounds.Right),
                    directBounds.Max(static bounds => bounds.Bottom));
            }
            double bottom = region.Lines.Min(static line => line.BaselineY - Math.Max(1D, line.FontSize * 0.25D));
            double top = region.Lines.Max(static line => line.BaselineY + Math.Max(1D, line.FontSize));
            return context.ToVisualBounds(region.XStart, bottom, region.XEnd, top);
        }

    }

    private sealed class BaselineGroup {
        internal BaselineGroup(double angle, double normal) { Angle = angle; Normal = normal; }
        internal double Angle { get; }
        internal double Normal { get; set; }
        internal List<PdfUnderstandingWord> Words { get; } = new();
    }

    private static double WordAnchorX(PdfUnderstandingWord word) => (word.XStart + word.XEnd) / 2D;
    internal static double NormalizeAngle(double value) { value %= 360D; if (value > 180D) value -= 360D; if (value <= -180D) value += 360D; return value; }
    private static double AngularDistance(double left, double right) { double distance = Math.Abs(NormalizeAngle(left) - NormalizeAngle(right)); return Math.Min(distance, 360D - distance); }
}
