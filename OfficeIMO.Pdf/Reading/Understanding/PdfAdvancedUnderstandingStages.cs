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
    /// <summary>Spatial connected-region segmentation.</summary>
    public static IPdfPageSegmentationStage PageSegmentation { get; } = new AdvancedPageSegmentationStage();
    /// <summary>Spanning-band and multi-column reading order.</summary>
    public static IPdfReadingOrderStage ReadingOrder { get; } = new PdfRecursiveXyCutReadingOrderStage();
    /// <summary>Business-document semantic classification.</summary>
    public static IPdfSemanticClassificationStage SemanticClassification { get; } = new AdvancedSemanticClassificationStage();

    private sealed class AdvancedGlyphDecodingStage : IPdfGlyphDecodingStage {
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            context.ThrowIfCancellationRequested();
            return context.Page.GetTextSpans();
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
            foreach (PdfUnderstandingWord word in words.OrderBy(static word => NormalizeAngle(word.RotationDegrees)).ThenByDescending(static word => word.BaselineY).ThenBy(static word => word.XStart)) {
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
                PdfUnderstandingWord[] ordered = group.Words.OrderBy(word => (Math.Cos(radians) * WordAnchorX(word)) + (Math.Sin(radians) * word.BaselineY)).ToArray();
                var runs = new List<List<PdfUnderstandingWord>> { new List<PdfUnderstandingWord>() };
                double previousAlong = double.NegativeInfinity;
                for (int i = 0; i < ordered.Length; i++) {
                    double along = (Math.Cos(radians) * WordAnchorX(ordered[i])) + (Math.Sin(radians) * ordered[i].BaselineY);
                    double splitGap = Math.Max(context.LayoutOptions.MinGutterWidth, ordered[i].FontSize * (Math.Abs(group.Angle) > 2D ? 6D : 5D));
                    if (runs[runs.Count - 1].Count > 0 && along - previousAlong > splitGap) runs.Add(new List<PdfUnderstandingWord>());
                    runs[runs.Count - 1].Add(ordered[i]);
                    previousAlong = along;
                }
                foreach (List<PdfUnderstandingWord> run in runs) {
                    PdfUnderstandingWord[] runWords = run.ToArray();
                    double normalSpread = runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Max() -
                        runWords.Select(word => (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY)).DefaultIfEmpty().Min();
                    lines.Add(new PdfUnderstandingLine(runWords, PdfInference.Clamp(runWords.Average(static word => word.Confidence) - Math.Min(0.25D, normalSpread / 20D)), new[] {
                        new PdfInferenceEvidence("line.arbitrary-baseline", "Words share a projected baseline at " + group.Angle.ToString("0.###", CultureInfo.InvariantCulture) + " degrees with " + normalSpread.ToString("0.###", CultureInfo.InvariantCulture) + " point spread.", normalSpread <= 2D ? 0.9D : 0.3D)
                    }));
                }
            }
            lines.Sort(static (left, right) => { int top = right.BaselineY.CompareTo(left.BaselineY); return top != 0 ? top : left.XStart.CompareTo(right.XStart); });
            return lines.Count == 0 ? Array.Empty<PdfUnderstandingLine>() : lines.AsReadOnly();
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
            var remaining = new HashSet<int>(Enumerable.Range(0, lines.Count));
            var regions = new List<PdfUnderstandingRegion>();
            AddCanonicalTableRegions(context, lines, remaining, regions);
            int[] parent = Enumerable.Range(0, lines.Count).ToArray();
            double maximumVerticalGap = lines.Count == 0 ? 0D : lines.Max(static line => line.FontSize) * 2.2D;
            int[] candidates = remaining.OrderBy(static index => index).ToArray();
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
                PdfUnderstandingLine[] ordered = component.Select(index => lines[index]).OrderByDescending(static line => line.BaselineY).ThenBy(static line => line.XStart).ToArray();
                double confidence = PdfInference.Clamp(ordered.Average(static line => line.Confidence) - Math.Min(0.2D, Math.Max(0, ordered.Length - 12) * 0.01D));
                regions.Add(new PdfUnderstandingRegion(ordered, confidence, new[] {
                    new PdfInferenceEvidence("region.spatial-connectivity", "The region is a connected component of " + ordered.Length.ToString(CultureInfo.InvariantCulture) + " line(s), allowing non-rectangular and mixed-layout neighborhoods.", ordered.Length > 1 ? 0.8D : 0.4D)
                }));
            }
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
        }

        private static void AddCanonicalTableRegions(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingLine> lines,
            HashSet<int> remaining,
            List<PdfUnderstandingRegion> regions) {
            if (context.DecodedRuns.Count == 0 || lines.Count == 0) return;

            StructuredPage structure = ContentStructureExtractor.Extract(
                context.DecodedRuns,
                context.LayoutOptions.ToEngineOptions(),
                context.Height);
            foreach (StructuredTable table in structure.TablesDetailed) {
                context.ConsumeWork();
                if (string.Equals(table.Kind, "leaders", StringComparison.OrdinalIgnoreCase) || table.Columns.Count < 2) continue;

                double top = Math.Max(table.YTop, table.YBottom);
                double bottom = Math.Min(table.YTop, table.YBottom);
                double left = table.Columns.Min(static column => Math.Min(column.From, column.To));
                double right = table.Columns.Max(static column => Math.Max(column.From, column.To));
                int[] tableLineIndexes = remaining
                    .Where(index => HasMeaningfulTableOverlap(lines[index], top, bottom, left, right))
                    .OrderByDescending(index => lines[index].BaselineY)
                    .ThenBy(index => lines[index].XStart)
                    .ToArray();
                if (tableLineIndexes.Length < 2) continue;

                PdfUnderstandingLine[] tableLines = tableLineIndexes.Select(index => lines[index]).ToArray();
                foreach (int index in tableLineIndexes) remaining.Remove(index);
                double confidence = PdfInference.Clamp(tableLines.Average(static line => line.Confidence));
                regions.Add(new PdfUnderstandingRegion(tableLines, confidence, new[] {
                    new PdfInferenceEvidence(
                        "region.canonical-table",
                        "The canonical layout engine recovered these aligned lines as one validated table.",
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
            if (region.YBottom <= context.Height * 0.08D && median > 0D && largest <= median * 0.9D) return (PdfUnderstandingSemanticKind.Footnote, 0.84D, "semantic.bottom-small-text", "Small text occupies the bottom eight percent of the page.");
            if (text.StartsWith("Figure ", StringComparison.OrdinalIgnoreCase) || text.StartsWith("Fig. ", StringComparison.OrdinalIgnoreCase) || text.StartsWith("Table ", StringComparison.OrdinalIgnoreCase)) return (PdfUnderstandingSemanticKind.Caption, 0.9D, "semantic.caption-prefix", "The region starts with a conventional figure or table caption prefix.");
            if (ContentStructureExtractor.IsListItemText(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, "semantic.list-marker", "The region begins with a bullet or numbered marker.");
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, 0.82D, "semantic.large-font", "The region font is materially larger than the page median.");
            return (PdfUnderstandingSemanticKind.Paragraph, 0.72D, "semantic.body-region", "No stronger business-document semantic signal was found.");
        }

    }

    private sealed class BaselineGroup {
        internal BaselineGroup(double angle, double normal) { Angle = angle; Normal = normal; }
        internal double Angle { get; }
        internal double Normal { get; set; }
        internal List<PdfUnderstandingWord> Words { get; } = new();
    }

    internal static bool HasMeaningfulTableOverlap(
        PdfUnderstandingLine line,
        double top,
        double bottom,
        double left,
        double right) {
        double verticalTolerance = Math.Max(1D, line.FontSize * 0.5D);
        if (line.BaselineY > top + verticalTolerance || line.BaselineY < bottom - verticalTolerance) return false;
        double lineLeft = Math.Min(line.XStart, line.XEnd);
        double lineRight = Math.Max(line.XStart, line.XEnd);
        double overlap = Math.Max(0D, Math.Min(lineRight, right) - Math.Max(lineLeft, left));
        double narrowerWidth = Math.Min(lineRight - lineLeft, right - left);
        return narrowerWidth > 0.001D && overlap + 0.001D >= narrowerWidth * 0.5D;
    }

    private static double WordAnchorX(PdfUnderstandingWord word) => (word.XStart + word.XEnd) / 2D;
    private static double NormalizeAngle(double value) { value %= 360D; if (value > 180D) value -= 360D; if (value <= -180D) value += 360D; return value; }
    private static double AngularDistance(double left, double right) { double distance = Math.Abs(NormalizeAngle(left) - NormalizeAngle(right)); return Math.Min(distance, 360D - distance); }
}
