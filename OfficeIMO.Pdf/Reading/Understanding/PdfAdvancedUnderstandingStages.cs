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
    public static IPdfTableDetectionStage TableDetection { get; } = new AdvancedTableDetectionStage();
    /// <summary>Spatial connected-region segmentation.</summary>
    public static IPdfPageSegmentationStage PageSegmentation { get; } = new AdvancedPageSegmentationStage();
    /// <summary>Spanning-band and multi-column reading order.</summary>
    public static IPdfReadingOrderStage ReadingOrder { get; } = new PdfRecursiveXyCutReadingOrderStage();
    /// <summary>Business-document semantic classification.</summary>
    public static IPdfSemanticClassificationStage SemanticClassification { get; } = new AdvancedSemanticClassificationStage();

    private static T[] CopyAndSort<T>(
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
                double perCharacter = text.Length > 0 && run.Advance > 0D ? run.Advance / text.Length : run.FontSize * 0.55D;
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
                    double startDistance = start * perCharacter;
                    double endDistance = cursor * perCharacter;
                    double startX = run.X + alongX * startDistance;
                    double startY = run.Y + alongY * startDistance;
                    double endX = run.X + alongX * endDistance;
                    double confidence = Math.Abs(run.RotationDegrees) <= 0.5D ? 0.96D : 0.9D;
                    result.Add(new PdfUnderstandingWord(
                        text.Substring(start, cursor - start),
                        Math.Min(startX, endX),
                        Math.Max(startX, endX),
                        startY,
                        run.FontSize,
                        NormalizeAngle(run.RotationDegrees),
                        new[] { run },
                        confidence,
                        new[] { new PdfInferenceEvidence("word.baseline-projection", "Word geometry was projected along a " + run.RotationDegrees.ToString("0.###", CultureInfo.InvariantCulture) + " degree baseline.", Math.Abs(run.RotationDegrees) <= 0.5D ? 0.8D : 0.6D) }));
                }
            }
            return result.Count == 0 ? Array.Empty<PdfUnderstandingWord>() : result.AsReadOnly();
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
                PdfUnderstandingWord[] ordered = CopyAndSort(
                    context,
                    group.Words,
                    (left, right) => ProjectAlong(left, radians).CompareTo(ProjectAlong(right, radians)));
                var runs = new List<List<PdfUnderstandingWord>> { new List<PdfUnderstandingWord>() };
                double previousAlongEnd = double.NegativeInfinity;
                for (int i = 0; i < ordered.Length; i++) {
                    double alongStart = GetProjectedAlongStart(ordered[i], radians);
                    double alongEnd = GetProjectedAlongEnd(ordered[i], radians);
                    double splitGap = Math.Max(context.LayoutOptions.MinGutterWidth, ordered[i].FontSize * (Math.Abs(group.Angle) > 2D ? 6D : 5D));
                    if (runs[runs.Count - 1].Count > 0 && alongStart - previousAlongEnd > splitGap) runs.Add(new List<PdfUnderstandingWord>());
                    runs[runs.Count - 1].Add(ordered[i]);
                    previousAlongEnd = Math.Max(previousAlongEnd, alongEnd);
                }
                foreach (List<PdfUnderstandingWord> run in runs) {
                    PdfUnderstandingWord[] runWords = run.ToArray();
                    string lineText = BuildLineText(context, runWords, group.Angle);
                    double normalSpread = runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Max() -
                        runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Min();
                    lines.Add(new PdfUnderstandingLine(runWords, lineText, PdfInference.Clamp(runWords.Average(static word => word.Confidence) - Math.Min(0.25D, normalSpread / 20D)), new[] {
                        new PdfInferenceEvidence("line.arbitrary-baseline", "Words share a projected baseline at " + group.Angle.ToString("0.###", CultureInfo.InvariantCulture) + " degrees with " + normalSpread.ToString("0.###", CultureInfo.InvariantCulture) + " point spread.", normalSpread <= 2D ? 0.9D : 0.3D)
                    }));
                }
            }
            PdfUnderstandingLine[] sortedLines = CopyAndSort(
                context,
                lines,
                static (left, right) => {
                    int top = right.BaselineY.CompareTo(left.BaselineY);
                    return top != 0 ? top : left.XStart.CompareTo(right.XStart);
                });
            return sortedLines.Length == 0 ? Array.Empty<PdfUnderstandingLine>() : Array.AsReadOnly(sortedLines);
        }

        private static double ProjectAlong(PdfUnderstandingWord word, double radians) =>
            (Math.Cos(radians) * WordAnchorX(word)) + (Math.Sin(radians) * word.BaselineY);

        private static string BuildLineText(
            PdfUnderstandingPageContext context,
            PdfUnderstandingWord[] words,
            double angle) {
            context.ThrowIfCancellationRequested();
            if (words.Length == 0) return string.Empty;
            var text = new System.Text.StringBuilder(words.Sum(static word => word.Text.Length) + words.Length);
            text.Append(words[0].Text);
            for (int wordIndex = 1; wordIndex < words.Length; wordIndex++) {
                PdfUnderstandingWord previous = words[wordIndex - 1];
                PdfUnderstandingWord current = words[wordIndex];
                if (NeedsSyntheticSpace(previous, current, angle)) text.Append(' ');
                text.Append(current.Text);
            }
            return text.ToString();
        }

        private static bool NeedsSyntheticSpace(
            PdfUnderstandingWord previous,
            PdfUnderstandingWord current,
            double angle) {
            if (SharesSourceRun(previous.SourceRuns, current.SourceRuns)) return true;
            if (HasExplicitBoundarySpace(previous.SourceRuns, current.SourceRuns)) return true;
            double radians = angle * Math.PI / 180D;
            double previousEnd = GetProjectedAlongEnd(previous, radians);
            double currentStart = GetProjectedAlongStart(current, radians);
            double threshold = Math.Max(1D, Math.Min(previous.FontSize, current.FontSize) * 0.18D);
            return currentStart - previousEnd > threshold;
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
            if (word.SourceRuns.Count != 1 || word.Text.Length == 0) return false;
            PdfTextSpan source = word.SourceRuns[0];
            string sourceText = source.Text ?? string.Empty;
            if (sourceText.Length == 0) return false;
            double perCharacter = source.Advance > 0D ? source.Advance / sourceText.Length : word.FontSize * 0.55D;
            advance = perCharacter * word.Text.Length;
            return advance > 0.001D;
        }

        private static double GetFallbackProjectedHalfExtent(PdfUnderstandingWord word, double radians) {
            double horizontalProjection = Math.Abs(Math.Cos(radians)) * Math.Max(0D, word.XEnd - word.XStart) / 2D;
            if (horizontalProjection > 0.001D) return horizontalProjection;
            return Math.Max(word.FontSize * 0.25D, word.FontSize * word.Text.Length * 0.2D);
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

    private sealed class AdvancedTableDetectionStage : IPdfTableDetectionStage {
        public IReadOnlyList<PdfUnderstandingTableCandidate> DetectTables(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingLine> lines) {
            if (context.DecodedRuns.Count == 0 || lines.Count == 0) {
                return Array.Empty<PdfUnderstandingTableCandidate>();
            }

            StructuredPage structure = ContentStructureExtractor.Extract(
                context.DecodedRuns,
                context.LayoutOptions.ToEngineOptions(),
                context.Height,
                context.ConsumeWork,
                context.ThrowIfCancellationRequested);
            var result = new List<PdfUnderstandingTableCandidate>(structure.TablesDetailed.Count);
            foreach (StructuredTable table in structure.TablesDetailed) {
                context.ConsumeWork();
                if (table.Columns.Count < 2 || table.SourceRuns.Count == 0) continue;

                var ownedRuns = new HashSet<PdfTextSpan>();
                for (int runIndex = 0; runIndex < table.SourceRuns.Count; runIndex++) {
                    context.ConsumeWork();
                    ownedRuns.Add(table.SourceRuns[runIndex]);
                }
                var matchedLines = new List<PdfUnderstandingLine>();
                for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                    context.ConsumeWork();
                    PdfUnderstandingLine line = lines[lineIndex];
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
                        matchedLines.Add(new PdfUnderstandingLine(
                            ownedWords.AsReadOnly(),
                            line.Confidence,
                            line.Evidence));
                    }
                }
                PdfUnderstandingLine[] sourceLines = CopyAndSort(
                    context,
                    matchedLines,
                    static (first, second) => {
                        int baseline = second.BaselineY.CompareTo(first.BaselineY);
                        return baseline != 0 ? baseline : first.XStart.CompareTo(second.XStart);
                    });
                if (sourceLines.Length < 2) continue;

                double confidence = PdfInference.Clamp(sourceLines.Average(static line => line.Confidence));
                result.Add(PdfUnderstandingTableCandidate.FromStructured(
                    table,
                    sourceLines,
                    confidence,
                    new[] {
                        new PdfInferenceEvidence(
                            "table.aligned-geometry",
                            "Repeated column geometry and row alignment form a bounded table candidate.",
                            0.9D)
                    },
                    context.ConsumeWork,
                    context.ThrowIfCancellationRequested));
            }
            return result.Count == 0 ? Array.Empty<PdfUnderstandingTableCandidate>() : result.AsReadOnly();
        }
    }

    private sealed class AdvancedPageSegmentationStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            lines = BuildSegmentationLines(context, lines);
            var remaining = new HashSet<int>(Enumerable.Range(0, lines.Count));
            var regions = new List<PdfUnderstandingRegion>();
            AddCanonicalTableRegions(context, lines, remaining, regions);
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
                        int baseline = right.BaselineY.CompareTo(left.BaselineY);
                        return baseline != 0 ? baseline : left.XStart.CompareTo(right.XStart);
                    });
                double confidence = PdfInference.Clamp(ordered.Average(static line => line.Confidence) - Math.Min(0.2D, Math.Max(0, ordered.Length - 12) * 0.01D));
                regions.Add(new PdfUnderstandingRegion(ordered, confidence, new[] {
                    new PdfInferenceEvidence("region.spatial-connectivity", "The region is a connected component of " + ordered.Length.ToString(CultureInfo.InvariantCulture) + " line(s), allowing non-rectangular and mixed-layout neighborhoods.", ordered.Length > 1 ? 0.8D : 0.4D)
                }));
            }
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
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
            double median = sizes.Length == 0 ? 0D : sizes[sizes.Length / 2];
            var result = new List<PdfUnderstandingSemanticElement>(orderedRegions.Count);
            foreach (PdfUnderstandingRegion region in orderedRegions) {
                context.ConsumeWork();
                (PdfUnderstandingSemanticKind kind, double confidence, string code, string message) = Classify(context, region, median);
                result.Add(new PdfUnderstandingSemanticElement(region, kind, confidence, new[] { new PdfInferenceEvidence(code, message, confidence - 0.5D) }));
            }
            return result.AsReadOnly();
        }

        private static (PdfUnderstandingSemanticKind Kind, double Confidence, string Code, string Message) Classify(PdfUnderstandingPageContext context, PdfUnderstandingRegion region, double median) {
            string text = region.Text.Trim();
            double largest = region.Lines.Max(static line => line.FontSize);
            if (region.Evidence.Any(static evidence => string.Equals(evidence.Code, "region.canonical-table", StringComparison.Ordinal))) return (PdfUnderstandingSemanticKind.Table, 0.93D, "semantic.canonical-table", "The canonical layout engine recovered the region as a validated table.");
            if (HasConventionalCaptionPrefix(text)) return (PdfUnderstandingSemanticKind.Caption, 0.9D, "semantic.caption-prefix", "The region starts with a conventional labeled figure or table caption prefix.");
            if (ContentStructureExtractor.IsListItemText(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, "semantic.list-marker", "The region begins with a bullet or numbered marker.");
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, 0.82D, "semantic.large-font", "The region font is materially larger than the page median.");
            if (IsCenteredAtVisualPageTop(context, region)) return (PdfUnderstandingSemanticKind.Heading, 0.76D, "semantic.centered-page-title", "Centered text occupies the top eight percent of the page and is treated as a page title.");
            if (context.Page.GetRotationDegrees() == 0 &&
                IsAtVisualPageTop(context, region) &&
                IsCompactRunningEdge(context, region, rejectLowercaseStart: true, rejectSentenceEnd: false)) return (PdfUnderstandingSemanticKind.Header, 0.78D, "semantic.page-edge-header", "Compact text occupies the top eight percent of the page and has no stronger semantic signal.");
            if (IsAtVisualPageBottom(context, region)) {
                if (median > 0D && largest <= median * 0.9D) return (PdfUnderstandingSemanticKind.Footnote, 0.84D, "semantic.bottom-small-text", "Small text occupies the bottom eight percent of the page.");
                if (context.Page.GetRotationDegrees() == 0 &&
                    IsCompactRunningEdge(context, region, rejectLowercaseStart: false, rejectSentenceEnd: true)) return (PdfUnderstandingSemanticKind.Footer, 0.78D, "semantic.page-edge-footer", "Compact text occupies the bottom eight percent of the page and has no stronger semantic signal.");
            }
            return (PdfUnderstandingSemanticKind.Paragraph, 0.72D, "semantic.body-region", "No stronger business-document semantic signal was found.");
        }

        private static bool HasConventionalCaptionPrefix(string text) {
            int labelStart;
            if (text.StartsWith("Figure ", StringComparison.OrdinalIgnoreCase)) {
                labelStart = "Figure ".Length;
            } else if (text.StartsWith("Fig. ", StringComparison.OrdinalIgnoreCase)) {
                labelStart = "Fig. ".Length;
            } else if (text.StartsWith("Table ", StringComparison.OrdinalIgnoreCase)) {
                labelStart = "Table ".Length;
            } else {
                return false;
            }

            int labelEnd = labelStart;
            while (labelEnd < text.Length &&
                   text[labelEnd] != ' ' &&
                   text[labelEnd] != '\t' &&
                   text[labelEnd] != ':' &&
                   text[labelEnd] != ';') {
                labelEnd++;
            }
            string label = text.Substring(labelStart, labelEnd - labelStart)
                .TrimEnd('.', ',', ':', ';', '-');
            if (label.Length == 0) return false;
            if (label.Any(static character => char.IsDigit(character))) return true;
            if (label.Length == 1 && char.IsLetter(label[0])) return true;
            return label.All(static character => char.ToUpperInvariant(character) is 'I' or 'V' or 'X' or 'L' or 'C' or 'D' or 'M');
        }

        private static bool IsCompactRunningEdge(
            PdfUnderstandingPageContext context,
            PdfUnderstandingRegion region,
            bool rejectLowercaseStart,
            bool rejectSentenceEnd) {
            if (region.Lines.Count != 1) return false;
            string text = region.Text.Trim();
            if (text.Length == 0 || text[text.Length - 1] == '-') return false;
            if (rejectLowercaseStart && char.IsLower(text[0])) return false;
            if (rejectSentenceEnd && text[text.Length - 1] is '.' or '?' or '!') return false;
            if (text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length > 4) return false;
            PdfVisualBounds visual = GetVisualBounds(context, region);
            double visualWidth = context.Page.GetVisualPageSize().Width;
            return visual.Right - visual.Left <= visualWidth * 0.45D;
        }

        private static bool IsCenteredAtVisualPageTop(PdfUnderstandingPageContext context, PdfUnderstandingRegion region) {
            string text = region.Text.Trim();
            if (context.Page.GetRotationDegrees() != 0 ||
                region.Lines.Count != 1 ||
                text.Length == 0 ||
                !char.IsUpper(text[0])) return false;
            PdfVisualBounds visual = GetVisualBounds(context, region);
            (double visualWidth, double visualHeight) = context.Page.GetVisualPageSize();
            double center = (visual.Left + visual.Right) / 2D;
            return visual.Top <= visualHeight * 0.08D &&
                Math.Abs(center - (visualWidth / 2D)) <= visualWidth * 0.12D;
        }

        private static bool IsAtVisualPageTop(PdfUnderstandingPageContext context, PdfUnderstandingRegion region) {
            PdfVisualBounds visual = GetVisualBounds(context, region);
            double visualHeight = context.Page.GetVisualPageSize().Height;
            return visual.Top <= visualHeight * 0.08D;
        }

        private static bool IsAtVisualPageBottom(PdfUnderstandingPageContext context, PdfUnderstandingRegion region) {
            PdfVisualBounds visual = GetVisualBounds(context, region);
            double visualHeight = context.Page.GetVisualPageSize().Height;
            return visual.Bottom >= visualHeight * 0.92D;
        }

        private static PdfVisualBounds GetVisualBounds(PdfUnderstandingPageContext context, PdfUnderstandingRegion region) {
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
    private static double NormalizeAngle(double value) { value %= 360D; if (value > 180D) value -= 360D; if (value <= -180D) value += 360D; return value; }
    private static double AngularDistance(double left, double right) { double distance = Math.Abs(NormalizeAngle(left) - NormalizeAngle(right)); return Math.Min(distance, 360D - distance); }
}
