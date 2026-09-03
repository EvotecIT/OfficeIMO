using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Built-in dependency-free understanding stages for rotated baselines, spatial regions, multi-column reading order, and richer business-document semantics.
/// </summary>
public static class PdfAdvancedUnderstandingStages {
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

    private sealed class AdvancedGlyphDecodingStage : IPdfGlyphDecodingStage {
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            context.ThrowIfCancellationRequested();
            return context.Page.GetTextSpans(context.CancellationToken);
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
                    result.Add(new PdfUnderstandingWord(
                        text.Substring(start, cursor - start),
                        Math.Min(startX, endX),
                        Math.Max(startX, endX),
                        startY,
                        run.FontSize,
                        NormalizeAngle(run.RotationDegrees),
                        new[] { run },
                        confidence,
                        new[] { new PdfInferenceEvidence(
                            hasResolvedCharacterAdvances ? "word.character-advance-projection" : "word.baseline-projection",
                            hasResolvedCharacterAdvances
                                ? "Word geometry was projected from decoded per-character advances along a " + run.RotationDegrees.ToString("0.###", CultureInfo.InvariantCulture) + " degree baseline."
                                : "Word geometry was projected uniformly along a " + run.RotationDegrees.ToString("0.###", CultureInfo.InvariantCulture) + " degree baseline because per-character advances were unavailable.",
                            hasResolvedCharacterAdvances ? 0.95D : (Math.Abs(run.RotationDegrees) <= 0.5D ? 0.8D : 0.6D)) },
                        Math.Max(0D, endDistance - startDistance),
                        visualBounds: null,
                        sourceSequence: sourceSequence));
                }
            }
            return result.Count == 0 ? Array.Empty<PdfUnderstandingWord>() : result.AsReadOnly();
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

    private sealed class AdvancedLineGroupingStage : IPdfLineGroupingStage {
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingWord> words) {
            var groups = new List<BaselineGroup>();
            var spatialIndex = new Dictionary<(int Angle, int Normal), List<BaselineGroup>>();
            PdfUnderstandingWord[] sortedWords = CopyAndSort(
                context,
                words,
                static (left, right) => {
                    int angle = NormalizeAngle(left.RotationDegrees).CompareTo(NormalizeAngle(right.RotationDegrees));
                    if (angle != 0) return angle;
                    int baseline = right.BaselineY.CompareTo(left.BaselineY);
                    return baseline != 0 ? baseline : left.XStart.CompareTo(right.XStart);
                });
            foreach (PdfUnderstandingWord word in sortedWords) {
                context.ConsumeWork();
                double angle = NormalizeAngle(word.RotationDegrees);
                double radians = angle * Math.PI / 180D;
                double normal = (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY);
                double tolerance = Math.Max(0.75D, Math.Min(context.LayoutOptions.LineMergeMaxPoints, word.FontSize * context.LayoutOptions.LineMergeToleranceEm));
                BaselineGroup? group = FindIndexedGroup(context, spatialIndex, angle, normal, tolerance);
                (int Angle, int Normal) previousKey = default;
                if (group is null) {
                    group = new BaselineGroup(angle, normal);
                    groups.Add(group);
                    AddToIndex(spatialIndex, group);
                } else {
                    previousKey = IndexKey(group.Angle, group.Normal);
                }
                group.Words.Add(word);
                group.Normal = ((group.Normal * (group.Words.Count - 1)) + normal) / group.Words.Count;
                if (group.Words.Count > 1) MoveInIndex(spatialIndex, group, previousKey);
            }

            var lines = new List<PdfUnderstandingLine>(groups.Count);
            foreach (BaselineGroup group in groups) {
                double radians = group.Angle * Math.PI / 180D;
                PdfUnderstandingWord[] sourceOrdered = CopyAndSort(
                    context,
                    group.Words,
                    static (left, right) => Nullable.Compare(left.SourceSequence, right.SourceSequence));
                PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
                    context.LayoutOptions.ReadingDirection,
                    sourceOrdered.Select(static word => word.Text));
                PdfUnderstandingWord[] ordered = CopyAndSort(
                    context,
                    group.Words,
                    (left, right) => direction == PdfReadingDirection.RightToLeft
                        ? CompareRightToLeft(left, right, radians)
                        : ProjectAlong(left, radians).CompareTo(ProjectAlong(right, radians)));
                var runs = new List<List<PdfUnderstandingWord>> { new List<PdfUnderstandingWord>() };
                double previousAlongEnd = double.NegativeInfinity;
                double previousAlongStart = double.PositiveInfinity;
                for (int i = 0; i < ordered.Length; i++) {
                    double alongStart = GetProjectedAlongStart(ordered[i], radians);
                    double alongEnd = GetProjectedAlongEnd(ordered[i], radians);
                    double splitGap = Math.Max(context.LayoutOptions.MinGutterWidth, ordered[i].FontSize * (Math.Abs(group.Angle) > 2D ? 6D : 5D));
                    double gap = direction == PdfReadingDirection.RightToLeft
                        ? previousAlongStart - alongEnd
                        : alongStart - previousAlongEnd;
                    if (runs[runs.Count - 1].Count > 0 && gap > splitGap) runs.Add(new List<PdfUnderstandingWord>());
                    runs[runs.Count - 1].Add(ordered[i]);
                    previousAlongEnd = Math.Max(previousAlongEnd, alongEnd);
                    previousAlongStart = alongStart;
                }
                foreach (List<PdfUnderstandingWord> run in runs) {
                    PdfUnderstandingWord[] runWords = run.ToArray();
                    string lineText = BuildLineText(context, runWords, group.Angle, direction);
                    double normalSpread = runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Max() -
                        runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Min();
                    int? lineSourceSequence = runWords.Any(static word => word.SourceSequence.HasValue)
                        ? runWords.Where(static word => word.SourceSequence.HasValue).Min(static word => word.SourceSequence!.Value)
                        : null;
                    lines.Add(new PdfUnderstandingLine(runWords, lineText, PdfInference.Clamp(runWords.Average(static word => word.Confidence) - Math.Min(0.25D, normalSpread / 20D)), new[] {
                        new PdfInferenceEvidence("line.arbitrary-baseline", "Words share a projected baseline at " + group.Angle.ToString("0.###", CultureInfo.InvariantCulture) + " degrees with " + normalSpread.ToString("0.###", CultureInfo.InvariantCulture) + " point spread.", normalSpread <= 2D ? 0.9D : 0.3D)
                    },
                    sourceSequence: lineSourceSequence));
                }
            }
            PdfReadingDirection pageDirection = PdfTextDirectionAnalysis.Resolve(
                context.LayoutOptions.ReadingDirection,
                words.OrderBy(static word => word.SourceSequence)
                    .Select(static word => word.Text));
            PdfUnderstandingLine[] sortedLines = CopyAndSort(
                context,
                lines,
                (left, right) => {
                    int top = right.BaselineY.CompareTo(left.BaselineY);
                    return top != 0
                        ? top
                        : pageDirection == PdfReadingDirection.RightToLeft
                            ? right.XStart.CompareTo(left.XStart)
                            : left.XStart.CompareTo(right.XStart);
                });
            return sortedLines.Length == 0 ? Array.Empty<PdfUnderstandingLine>() : Array.AsReadOnly(sortedLines);
        }

        private static double ProjectAlong(PdfUnderstandingWord word, double radians) =>
            (Math.Cos(radians) * WordAnchorX(word)) + (Math.Sin(radians) * word.BaselineY);

        private static int CompareRightToLeft(
            PdfUnderstandingWord left,
            PdfUnderstandingWord right,
            double radians) {
            if (SharesSourceRun(left.SourceRuns, right.SourceRuns) &&
                left.SourceSequence.HasValue &&
                right.SourceSequence.HasValue) {
                int sourceOrder = left.SourceSequence.Value.CompareTo(right.SourceSequence.Value);
                if (sourceOrder != 0) return sourceOrder;
            }
            int geometry = ProjectAlong(right, radians).CompareTo(ProjectAlong(left, radians));
            return geometry != 0
                ? geometry
                : Nullable.Compare(left.SourceSequence, right.SourceSequence);
        }

        private static string BuildLineText(
            PdfUnderstandingPageContext context,
            PdfUnderstandingWord[] words,
            double angle,
            PdfReadingDirection direction) {
            context.ThrowIfCancellationRequested();
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
            if (HasExplicitBoundarySpace(previous.SourceRuns, current.SourceRuns)) return true;
            double radians = angle * Math.PI / 180D;
            double gap = direction == PdfReadingDirection.RightToLeft
                ? GetProjectedAlongStart(previous, radians) - GetProjectedAlongEnd(current, radians)
                : GetProjectedAlongStart(current, radians) - GetProjectedAlongEnd(previous, radians);
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
            IReadOnlyList<PdfTextSpan> right) {
            if (left.Count == 0 || right.Count == 0) return false;
            PdfTextSpan previous = left[left.Count - 1];
            PdfTextSpan current = right[0];
            return previous.LogicalTrailingSpace ||
                   current.LogicalLeadingSpace ||
                   (previous.Text.Length > 0 && char.IsWhiteSpace(previous.Text[previous.Text.Length - 1])) ||
                   (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[0]));
        }

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

        private static BaselineGroup? FindIndexedGroup(
            PdfUnderstandingPageContext context,
            Dictionary<(int Angle, int Normal), List<BaselineGroup>> index,
            double angle,
            double normal,
            double tolerance) {
            (int angleBucket, int normalBucket) = IndexKey(angle, normal);
            int normalRadius = (int)Math.Ceiling(tolerance / 0.75D) + 1;
            for (int angleOffset = -2; angleOffset <= 2; angleOffset++) {
                int candidateAngle = (angleBucket + angleOffset + 180) % 180;
                for (int normalOffset = -normalRadius; normalOffset <= normalRadius; normalOffset++) {
                    context.ConsumeWork();
                    if (!index.TryGetValue((candidateAngle, normalBucket + normalOffset), out List<BaselineGroup>? candidates)) continue;
                    for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                        context.ConsumeWork();
                        BaselineGroup candidate = candidates[candidateIndex];
                        if (AngularDistance(candidate.Angle, angle) <= 2D && Math.Abs(candidate.Normal - normal) <= tolerance) return candidate;
                    }
                }
            }
            return null;
        }

        private static void AddToIndex(Dictionary<(int Angle, int Normal), List<BaselineGroup>> index, BaselineGroup group) {
            (int Angle, int Normal) key = IndexKey(group.Angle, group.Normal);
            if (!index.TryGetValue(key, out List<BaselineGroup>? values)) {
                values = new List<BaselineGroup>();
                index.Add(key, values);
            }
            values.Add(group);
        }

        private static void MoveInIndex(
            Dictionary<(int Angle, int Normal), List<BaselineGroup>> index,
            BaselineGroup group,
            (int Angle, int Normal) previousKey) {
            (int Angle, int Normal) nextKey = IndexKey(group.Angle, group.Normal);
            if (nextKey == previousKey) return;
            if (index.TryGetValue(previousKey, out List<BaselineGroup>? previous)) {
                previous.Remove(group);
                if (previous.Count == 0) index.Remove(previousKey);
            }
            AddToIndex(index, group);
        }

        private static (int Angle, int Normal) IndexKey(double angle, double normal) {
            int angleBucket = ((int)Math.Floor((NormalizeAngle(angle) + 180D) / 2D)) % 180;
            if (angleBucket < 0) angleBucket += 180;
            return (angleBucket, (int)Math.Floor(normal / 0.75D));
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
                AddStructuralSubregions(context, ordered, pageMedianFontSize, regions);
            }
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
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
                    "region.structural-connectivity",
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
                    result.Add(new PdfUnderstandingLine(
                        unownedWords.AsReadOnly(),
                        line.Confidence,
                        line.Evidence));
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
                        "region.canonical-table",
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
            double[] sizes = orderedRegions.SelectMany(static region => region.Lines).Select(static line => line.FontSize).OrderBy(static size => size).ToArray();
            double median = sizes.Length == 0 ? 0D : sizes[(sizes.Length - 1) / 2];
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
            if (region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region.canonical-table", StringComparison.Ordinal))) return (PdfUnderstandingSemanticKind.Table, 0.93D, "semantic.canonical-table", "The canonical layout engine recovered the region as a validated table.");
            if (IsAdjacentTableCaption(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Caption, 0.78D, "semantic.table-caption-geometry", "A compact line is immediately adjacent to and aligned with a validated table.");
            if (ContentStructureExtractor.IsListItemText(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, "semantic.list-marker", "The region begins with a bullet or numbered marker.");
            if (IsLocalTypographicLead(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Heading, 0.8D, "semantic.local-font-transition", "A compact line is materially larger than the nearby aligned content it introduces.");
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, 0.82D, "semantic.large-font", "The region font is materially larger than the page median.");
            if (IsCenteredSpanningBand(context, region, pageRegions)) return (PdfUnderstandingSemanticKind.Heading, 0.76D, "semantic.centered-spanning-band", "A compact centered line forms a band above or below two separated content columns.");
            return (PdfUnderstandingSemanticKind.Paragraph, 0.72D, "semantic.body-region", "No stronger business-document semantic signal was found.");
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
                    !table.Evidence.Any(static evidence => string.Equals(evidence.Code, "region.canonical-table", StringComparison.Ordinal))) continue;

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
                    region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region.canonical-table", StringComparison.Ordinal))) continue;
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
            (double pageWidth, _) = context.Page.GetVisualPageSize();
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
            return context.Page.TransformBoundsToVisual(region.XStart, bottom, region.XEnd, top);
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
