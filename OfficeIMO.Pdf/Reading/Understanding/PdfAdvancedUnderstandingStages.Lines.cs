using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfAdvancedUnderstandingStages {
    private sealed class AdvancedLineGroupingStage : IPdfLineGroupingStage {
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingWord> words) {
            var groups = new List<BaselineGroup>();
            var spatialIndex = new BaselineIndex();
            var angledWords = new (double Angle, PdfUnderstandingWord Word)[words.Count];
            for (int index = 0; index < angledWords.Length; index++) angledWords[index] = (NormalizeAngle(words[index].RotationDegrees), words[index]);
            (double Angle, PdfUnderstandingWord Word)[] sortedWords = CopyAndSort(
                context,
                angledWords,
                static (left, right) => {
                    int angle = left.Angle.CompareTo(right.Angle);
                    if (angle != 0) return angle;
                    int baseline = right.Word.BaselineY.CompareTo(left.Word.BaselineY);
                    return baseline != 0 ? baseline : left.Word.XStart.CompareTo(right.Word.XStart);
                });
            foreach ((double angle, PdfUnderstandingWord word) in sortedWords) {
                context.ConsumeWork();
                double radians = angle * Math.PI / 180D;
                bool selectionAnchor = IsSelectionAnchor(word);
                double normal = GetGroupingNormal(word, radians);
                double tolerance = Math.Max(0.75D, Math.Min(context.LayoutOptions.LineMergeMax.ToPoints(), word.FontSize * context.LayoutOptions.LineMergeToleranceEm));
                if (selectionAnchor) tolerance = Math.Max(tolerance,
                    Math.Min(context.LayoutOptions.LineMergeMax.ToPoints() * 2D, word.FontSize * 0.2D));
                BaselineGroup? group = FindIndexedGroup(context, spatialIndex, angle, normal, tolerance, selectionAnchor);
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

            var gutterAnchors = CollectGutterAnchors(context, groups);
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
                PdfUnderstandingWord[] ordered = GetAlongOrdered(context, group);
                var runs = new List<List<PdfUnderstandingWord>> { new List<PdfUnderstandingWord>() };
                double previousAlongEnd = double.NegativeInfinity;
                for (int i = 0; i < ordered.Length; i++) {
                    double alongStart = GetProjectedAlongStart(ordered[i], radians);
                    double alongEnd = GetProjectedAlongEnd(ordered[i], radians);
                    double splitGap = Math.Max(context.LayoutOptions.MinGutterWidth.ToPoints(), ordered[i].FontSize * (Math.Abs(group.Angle) > 2D ? 6D : 5D));
                    double gap = alongStart - previousAlongEnd;
                    bool repeatedGutter = gap >= GetOrdinaryGutterWidth(context, ordered[i]) &&
                        HasRepeatedGutter(gutterAnchors.Standard, group, alongStart) ||
                        gap >= GetRepeatedGutterWidth(context, ordered[i]) && HasRepeatedGutter(gutterAnchors.Prose, group, alongStart) &&
                        HasProseOnBothSides(context, ordered, i, radians) &&
                        HasClearGutter(context, groups, group, alongStart, ordered[i].FontSize, radians);
                    bool hangingMarker = runs.Count == 1 && i == 1 && ContentStructureExtractor.IsHangingListMarker(ordered[0].Text, gap);
                    if (runs[runs.Count - 1].Count > 0 && (gap > splitGap || repeatedGutter) && !hangingMarker) runs.Add(new List<PdfUnderstandingWord>());
                    runs[runs.Count - 1].Add(ordered[i]);
                    previousAlongEnd = Math.Max(previousAlongEnd, alongEnd);
                }
                foreach (List<PdfUnderstandingWord> run in runs) {
                    PdfUnderstandingWord[] runWords = OrderLogicalWords(context, run, direction);
                    string lineText = ComposeLineText(context, runWords, group.Angle, direction);
                    double normalSpread = runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Max() -
                        runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Min();
                    int? lineSourceSequence = runWords.Any(static word => word.SourceSequence.HasValue)
                        ? runWords.Where(static word => word.SourceSequence.HasValue).Min(static word => word.SourceSequence!.Value)
                        : null;
                    lines.Add(new PdfUnderstandingLine(runWords, lineText, PdfInference.Clamp(runWords.Average(static word => word.Confidence) - Math.Min(0.25D, normalSpread / 20D)), new[] {
                        new PdfInferenceEvidence("line-arbitrary-baseline", "Words share a projected baseline at " + group.Angle.ToString("0.###", CultureInfo.InvariantCulture) + " degrees with " + normalSpread.ToString("0.###", CultureInfo.InvariantCulture) + " point spread.", normalSpread <= 2D ? 0.9D : 0.3D)
                    },
                    sourceSequence: lineSourceSequence,
                    visualBounds: runWords.All(static word => word.VisualBounds is not null)
                        ? new PdfLogicalVisualBounds(runWords.Min(static word => word.VisualBounds!.Left),
                            runWords.Min(static word => word.VisualBounds!.Top), runWords.Max(static word => word.VisualBounds!.Right),
                            runWords.Max(static word => word.VisualBounds!.Bottom)) : null));
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

        private static PdfUnderstandingWord[] OrderLogicalWords(PdfUnderstandingPageContext context,
            List<PdfUnderstandingWord> visualWords, PdfReadingDirection direction) {
            if (context.LayoutOptions.ReadingDirection == PdfReadingDirection.LeftToRight) return visualWords.ToArray();
            context.ConsumeWork(visualWords.Count);
            PdfUnderstandingWord[] ordered = PdfTextDirectionAnalysis.RestoreLogicalFragmentOrder(
                visualWords, static word => word.Text, direction, context.CancellationToken);
            // A decoded run may already carry a logical multiword replacement. Preserve that
            // run's internal sequence while reordering independent positioned fragments around it.
            for (int start = 0; start < ordered.Length;) {
                int end = start + 1;
                while (end < ordered.Length && SharesSourceRun(ordered[start].SourceRuns, ordered[end].SourceRuns)) end++;
                if (end - start > 1 && ordered.Skip(start).Take(end - start).All(static word => word.SourceSequence.HasValue)) {
                    PdfUnderstandingWord[] sourceOrdered = CopyAndSort(context, ordered.Skip(start).Take(end - start).ToArray(),
                        static (left, right) => Nullable.Compare(left.SourceSequence, right.SourceSequence));
                    sourceOrdered.CopyTo(ordered, start);
                }
                start = end;
            }
            return ordered;
        }

        private static double GetGroupingNormal(PdfUnderstandingWord word, double radians) {
            double normal = (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY);
            // Invisible ActualText anchors describe selection boxes, whose lower edges depend on
            // the recognized glyphs. Their centers are a better line-alignment signal. Keep the
            // original baseline and bounds on the word so extraction never changes selection geometry.
            if (IsSelectionAnchor(word))
                normal += word.FontSize * 0.5D;
            return normal;
        }

        private static bool IsSelectionAnchor(PdfUnderstandingWord word) =>
            word.IsSelectionBox || (word.SourceRuns.Count > 0 && word.SourceRuns.All(static run => run.HasActualText && run.TextRenderingMode == 3));

        private sealed class GutterAnchors {
            internal readonly Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)> Standard = new();
            internal readonly Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)> Prose = new();
        }

        private static GutterAnchors CollectGutterAnchors(
            PdfUnderstandingPageContext context, List<BaselineGroup> groups) {
            var anchors = new GutterAnchors();
            foreach (BaselineGroup group in groups) {
                double radians = group.Angle * Math.PI / 180D;
                PdfUnderstandingWord[] ordered = OrderAlong(context, group);
                double previousEnd = double.NegativeInfinity;
                for (int index = 0; index < ordered.Length; index++) {
                    context.ConsumeWork();
                    double start = GetProjectedAlongStart(ordered[index], radians);
                    double gap = start - previousEnd;
                    if (index > 0 && TryGutterKey(group, start, out var key)) {
                        if (gap >= GetOrdinaryGutterWidth(context, ordered[index])) Observe(anchors.Standard, group, key);
                        if (gap >= GetRepeatedGutterWidth(context, ordered[index]) &&
                            HasProseOnBothSides(context, ordered, index, radians) &&
                            HasClearGutter(context, groups, group, start, ordered[index].FontSize, radians)) {
                            Observe(anchors.Prose, group, key);
                        }
                    }
                    previousEnd = Math.Max(previousEnd, GetProjectedAlongEnd(ordered[index], radians));
                }
            }
            return anchors;

            static void Observe(Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)> map,
                BaselineGroup group, (int Angle, long Along) key) {
                if (map.TryGetValue(key, out var existing)) {
                    if (!ReferenceEquals(existing.First, group)) map[key] = (existing.First, true);
                } else map.Add(key, (group, false));
            }
        }

        // Repeated spaces inside justified prose are not columns. A narrow
        // gutter must remain free of words on nearby baselines as well.
        private static bool HasClearGutter(PdfUnderstandingPageContext context,
            List<BaselineGroup> groups, BaselineGroup candidate, double end, double fontSize, double radians) {
            double start = end - Math.Min(context.LayoutOptions.MinGutterWidth.ToPoints(), Math.Max(6D, fontSize * 0.8D));
            foreach (BaselineGroup group in groups) {
                context.ConsumeWork();
                if (Math.Abs(group.Angle - candidate.Angle) > 2D ||
                    Math.Abs(group.Normal - candidate.Normal) > fontSize * 6D) continue;
                foreach (PdfUnderstandingWord word in group.Words) {
                    context.ConsumeWork();
                    if (GetProjectedAlongEnd(word, radians) > start + 0.5D &&
                        GetProjectedAlongStart(word, radians) < end - 0.5D) return false;
                }
            }
            return true;
        }

        private static double GetOrdinaryGutterWidth(PdfUnderstandingPageContext context, PdfUnderstandingWord word) =>
            Math.Min(context.LayoutOptions.MinGutterWidth.ToPoints(), Math.Max(9D, word.FontSize * 1.5D));

        // A narrow repeated gap needs prose on either side. Aligned short table
        // cells also repeat their gaps, but splitting those changes row text.
        private static bool HasProseOnBothSides(PdfUnderstandingPageContext context,
            PdfUnderstandingWord[] words, int split, double radians) {
            if (split <= 0 || split >= words.Length) return false;
            double angle = radians * 180D / Math.PI;
            return Count(-1, split - 1) >= 3 && Count(1, split) >= 3;

            int Count(int step, int start) {
                // Positioned word artifacts can be individual glyphs. Use the
                // line composer's boundaries, counting alphabetic words only.
                int count = 0;
                bool hasLetters = PdfUnicodeScalarAnalysis.ContainsLetter(words[start].Text);
                for (int index = start + step; index >= 0 && index < words.Length && count < 3; index += step) {
                    context.ConsumeWork();
                    int left = Math.Min(index, index - step), right = left + 1;
                    double gap = GetProjectedAlongStart(words[right], radians) - GetProjectedAlongEnd(words[left], radians);
                    if (gap >= GetRepeatedGutterWidth(context, words[right])) break;
                    if (NeedsSyntheticSpace(words[left], words[right], angle, PdfReadingDirection.LeftToRight)) {
                        if (hasLetters) count++;
                        if (count >= 3) return count;
                        hasLetters = false;
                    }
                    hasLetters |= PdfUnicodeScalarAnalysis.ContainsLetter(words[index].Text);
                }
                return count + (hasLetters ? 1 : 0);
            }
        }

        /// <summary>
        /// A gap that recurs at the same position on several baselines is a column gutter even when it is narrower than
        /// the page-level gutter option. Narrower gaps additionally require prose on both sides.
        /// </summary>
        private static double GetRepeatedGutterWidth(PdfUnderstandingPageContext context, PdfUnderstandingWord word) =>
            Math.Min(context.LayoutOptions.MinGutterWidth.ToPoints(), Math.Max(6D, Math.Max(1D, word.FontSize) * 0.8D));

        private static bool HasRepeatedGutter(Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)> anchors,
            BaselineGroup group, double start) {
            if (!TryGutterKey(group, start, out var key)) return false;
            for (int offset = -1; offset <= 1; offset++) {
                if (anchors.TryGetValue((key.Angle, key.Along + offset), out var candidate) &&
                    (candidate.Repeated || !ReferenceEquals(candidate.First, group))) return true;
            }
            return false;
        }

        private static bool TryGutterKey(BaselineGroup group, double start, out (int Angle, long Along) key) {
            key = default;
            if (double.IsNaN(start) || double.IsInfinity(start) || Math.Abs(start) >= long.MaxValue - 2048D) return false;
            key = ((int)Math.Round(group.Angle), (long)Math.Round(start));
            return true;
        }

        // Orders a group's words along its baseline once. Line splitting needs the same order
        // after gutter collection; it reuses the array and charges the same work again.
        private static PdfUnderstandingWord[] OrderAlong(PdfUnderstandingPageContext context, BaselineGroup group) {
            double radians = group.Angle * Math.PI / 180D;
            var keyed = new (double Along, PdfUnderstandingWord Word)[group.Words.Count];
            for (int index = 0; index < keyed.Length; index++) keyed[index] = (ProjectAlong(group.Words[index], radians), group.Words[index]);
            (double Along, PdfUnderstandingWord Word)[] sorted = CopyAndSort(context, keyed,
                static (left, right) => left.Along.CompareTo(right.Along), out long work);
            var ordered = new PdfUnderstandingWord[sorted.Length];
            for (int index = 0; index < sorted.Length; index++) ordered[index] = sorted[index].Word;
            group.AlongOrdered = ordered;
            group.AlongOrderWork = work;
            return ordered;
        }

        private static PdfUnderstandingWord[] GetAlongOrdered(PdfUnderstandingPageContext context, BaselineGroup group) {
            if (group.AlongOrdered is not PdfUnderstandingWord[] ordered) return OrderAlong(context, group);
            if (group.AlongOrderWork > 0) context.ConsumeWork(group.AlongOrderWork);
            return ordered;
        }

        // Baseline groups bucketed by angle (two-degree buckets) and normal (0.75 pt buckets).
        // Each angle bucket remembers the range of normal buckets it has ever held, so lookups
        // skip angles with no groups and normals outside that range.
        private sealed class BaselineIndex {
            internal readonly Dictionary<(int Angle, int Normal), List<BaselineGroup>> Buckets = new();
            internal readonly bool[] Occupied = new bool[180];
            internal readonly int[] MinNormal = new int[180];
            internal readonly int[] MaxNormal = new int[180];
        }

        private static BaselineGroup? FindIndexedGroup(
            PdfUnderstandingPageContext context,
            BaselineIndex index,
            double angle,
            double normal,
            double tolerance, bool selectionAnchor) {
            (int angleBucket, int normalBucket) = IndexKey(angle, normal);
            int normalRadius = (int)Math.Ceiling(tolerance / 0.75D) + 1;
            int first = normalBucket - normalRadius, last = normalBucket + normalRadius;
            for (int angleOffset = -2; angleOffset <= 2; angleOffset++) {
                int candidateAngle = (angleBucket + angleOffset + 180) % 180;
                // Every probed bucket costs one work unit, including the ones skipped here as empty.
                int low = index.Occupied[candidateAngle] ? Math.Max(first, index.MinNormal[candidateAngle]) : last + 1;
                int high = index.Occupied[candidateAngle] ? Math.Min(last, index.MaxNormal[candidateAngle]) : last;
                if (low > high) {
                    context.ConsumeWork(last - first + 1);
                    continue;
                }
                if (low > first) context.ConsumeWork(low - first);
                for (int normalKey = low; normalKey <= high; normalKey++) {
                    context.ConsumeWork();
                    if (!index.Buckets.TryGetValue((candidateAngle, normalKey), out List<BaselineGroup>? candidates)) continue;
                    for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                        context.ConsumeWork();
                        BaselineGroup candidate = candidates[candidateIndex];
                        if (candidate.Words.Count > 0 && IsSelectionAnchor(candidate.Words[0]) != selectionAnchor) continue;
                        if (AngularDistance(candidate.Angle, angle) <= 2D && Math.Abs(candidate.Normal - normal) <= tolerance) return candidate;
                    }
                }
                if (high < last) context.ConsumeWork(last - high);
            }
            return null;
        }

        private static void AddToIndex(BaselineIndex index, BaselineGroup group) {
            (int Angle, int Normal) key = IndexKey(group.Angle, group.Normal);
            if (!index.Buckets.TryGetValue(key, out List<BaselineGroup>? values)) {
                values = new List<BaselineGroup>();
                index.Buckets.Add(key, values);
            }
            values.Add(group);
            if (!index.Occupied[key.Angle]) {
                index.Occupied[key.Angle] = true;
                index.MinNormal[key.Angle] = key.Normal;
                index.MaxNormal[key.Angle] = key.Normal;
            } else {
                index.MinNormal[key.Angle] = Math.Min(index.MinNormal[key.Angle], key.Normal);
                index.MaxNormal[key.Angle] = Math.Max(index.MaxNormal[key.Angle], key.Normal);
            }
        }

        private static void MoveInIndex(
            BaselineIndex index,
            BaselineGroup group,
            (int Angle, int Normal) previousKey) {
            (int Angle, int Normal) nextKey = IndexKey(group.Angle, group.Normal);
            if (nextKey == previousKey) return;
            if (index.Buckets.TryGetValue(previousKey, out List<BaselineGroup>? previous)) {
                previous.Remove(group);
                if (previous.Count == 0) index.Buckets.Remove(previousKey);
            }
            AddToIndex(index, group);
        }

        private static (int Angle, int Normal) IndexKey(double angle, double normal) {
            int angleBucket = ((int)Math.Floor((NormalizeAngle(angle) + 180D) / 2D)) % 180;
            if (angleBucket < 0) angleBucket += 180;
            return (angleBucket, (int)Math.Floor(normal / 0.75D));
        }
    }

}
