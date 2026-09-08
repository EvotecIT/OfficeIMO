using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfAdvancedUnderstandingStages {
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
                bool selectionAnchor = IsSelectionAnchor(word);
                double normal = GetGroupingNormal(word, radians);
                double tolerance = Math.Max(0.75D, Math.Min(context.LayoutOptions.LineMergeMaxPoints, word.FontSize * context.LayoutOptions.LineMergeToleranceEm));
                if (selectionAnchor) tolerance = Math.Max(tolerance,
                    Math.Min(context.LayoutOptions.LineMergeMaxPoints * 2D, word.FontSize * 0.2D));
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
                    double gap = alongStart - previousAlongEnd;
                    bool repeatedGutter = gap >= context.LayoutOptions.MinGutterWidth && HasRepeatedGutter(gutterAnchors, group, alongStart);
                    if (runs[runs.Count - 1].Count > 0 && (gap > splitGap || repeatedGutter)) runs.Add(new List<PdfUnderstandingWord>());
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
                        new PdfInferenceEvidence("line.arbitrary-baseline", "Words share a projected baseline at " + group.Angle.ToString("0.###", CultureInfo.InvariantCulture) + " degrees with " + normalSpread.ToString("0.###", CultureInfo.InvariantCulture) + " point spread.", normalSpread <= 2D ? 0.9D : 0.3D)
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

        private static Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)> CollectGutterAnchors(
            PdfUnderstandingPageContext context, List<BaselineGroup> groups) {
            var anchors = new Dictionary<(int Angle, long Along), (BaselineGroup First, bool Repeated)>();
            foreach (BaselineGroup group in groups) {
                double radians = group.Angle * Math.PI / 180D;
                PdfUnderstandingWord[] ordered = CopyAndSort(context, group.Words,
                    (left, right) => ProjectAlong(left, radians).CompareTo(ProjectAlong(right, radians)));
                double previousEnd = double.NegativeInfinity;
                for (int index = 0; index < ordered.Length; index++) {
                    context.ConsumeWork();
                    double start = GetProjectedAlongStart(ordered[index], radians);
                    if (index > 0 && start - previousEnd >= context.LayoutOptions.MinGutterWidth && TryGutterKey(group, start, out var key)) {
                        if (anchors.TryGetValue(key, out var existing)) {
                            if (!ReferenceEquals(existing.First, group)) anchors[key] = (existing.First, true);
                        } else anchors.Add(key, (group, false));
                    }
                    previousEnd = Math.Max(previousEnd, GetProjectedAlongEnd(ordered[index], radians));
                }
            }
            return anchors;
        }

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

        private static BaselineGroup? FindIndexedGroup(
            PdfUnderstandingPageContext context,
            Dictionary<(int Angle, int Normal), List<BaselineGroup>> index,
            double angle,
            double normal,
            double tolerance, bool selectionAnchor) {
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
                        if (candidate.Words.Count > 0 && IsSelectionAnchor(candidate.Words[0]) != selectionAnchor) continue;
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

}
