using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Built-in dependency-free understanding stages for rotated baselines, spatial regions, multi-column reading order, and richer business-document semantics.
/// </summary>
public static class PdfAdvancedUnderstandingStages {
    /// <summary>Rotation-aware word grouping.</summary>
    public static IPdfWordGroupingStage WordGrouping { get; } = new AdvancedWordGroupingStage();
    /// <summary>Arbitrary-baseline line grouping.</summary>
    public static IPdfLineGroupingStage LineGrouping { get; } = new AdvancedLineGroupingStage();
    /// <summary>Spatial connected-region segmentation.</summary>
    public static IPdfPageSegmentationStage PageSegmentation { get; } = new AdvancedPageSegmentationStage();
    /// <summary>Spanning-band and multi-column reading order.</summary>
    public static IPdfReadingOrderStage ReadingOrder { get; } = new AdvancedReadingOrderStage();
    /// <summary>Business-document semantic classification.</summary>
    public static IPdfSemanticClassificationStage SemanticClassification { get; } = new AdvancedSemanticClassificationStage();

    private sealed class AdvancedWordGroupingStage : IPdfWordGroupingStage {
        public IReadOnlyList<PdfUnderstandingWord> GroupWords(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
            var result = new List<PdfUnderstandingWord>();
            for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                PdfTextSpan run = runs[runIndex];
                string text = run.Text ?? string.Empty;
                double radians = run.RotationDegrees * Math.PI / 180D;
                double alongX = Math.Cos(radians);
                double alongY = Math.Sin(radians);
                double perCharacter = text.Length > 0 && run.Advance > 0D ? run.Advance / text.Length : run.FontSize * 0.55D;
                int cursor = 0;
                while (cursor < text.Length) {
                    while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) cursor++;
                    int start = cursor;
                    while (cursor < text.Length && !char.IsWhiteSpace(text[cursor])) cursor++;
                    if (cursor == start) continue;
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
            foreach (PdfUnderstandingWord word in words.OrderBy(static word => NormalizeAngle(word.RotationDegrees)).ThenByDescending(static word => word.BaselineY).ThenBy(static word => word.XStart)) {
                double angle = NormalizeAngle(word.RotationDegrees);
                double radians = angle * Math.PI / 180D;
                double normal = (-Math.Sin(radians) * WordAnchorX(word)) + (Math.Cos(radians) * word.BaselineY);
                double tolerance = Math.Max(0.75D, Math.Min(context.LayoutOptions.LineMergeMaxPoints, word.FontSize * context.LayoutOptions.LineMergeToleranceEm));
                BaselineGroup? group = groups.FirstOrDefault(candidate => AngularDistance(candidate.Angle, angle) <= 2D && Math.Abs(candidate.Normal - normal) <= tolerance);
                if (group is null) { group = new BaselineGroup(angle, normal); groups.Add(group); }
                group.Words.Add(word);
                group.Normal = ((group.Normal * (group.Words.Count - 1)) + normal) / group.Words.Count;
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
    }

    private sealed class AdvancedPageSegmentationStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            var remaining = new HashSet<int>(Enumerable.Range(0, lines.Count));
            var regions = new List<PdfUnderstandingRegion>();
            while (remaining.Count > 0) {
                int seed = remaining.First();
                remaining.Remove(seed);
                var component = new List<int> { seed };
                var queue = new Queue<int>(); queue.Enqueue(seed);
                while (queue.Count > 0) {
                    int current = queue.Dequeue();
                    foreach (int candidate in remaining.ToArray()) {
                        if (!AreSpatialNeighbors(lines[current], lines[candidate])) continue;
                        remaining.Remove(candidate); component.Add(candidate); queue.Enqueue(candidate);
                    }
                }
                PdfUnderstandingLine[] ordered = component.Select(index => lines[index]).OrderByDescending(static line => line.BaselineY).ThenBy(static line => line.XStart).ToArray();
                double confidence = PdfInference.Clamp(ordered.Average(static line => line.Confidence) - Math.Min(0.2D, Math.Max(0, ordered.Length - 12) * 0.01D));
                regions.Add(new PdfUnderstandingRegion(ordered, confidence, new[] {
                    new PdfInferenceEvidence("region.spatial-connectivity", "The region is a connected component of " + ordered.Length.ToString(CultureInfo.InvariantCulture) + " line(s), allowing non-rectangular and mixed-layout neighborhoods.", ordered.Length > 1 ? 0.8D : 0.4D)
                }));
            }
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
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

    private sealed class AdvancedReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) {
            if (regions.Count <= 1) return regions.ToArray();
            double pageWidth = Math.Max(1D, context.Width);
            PdfUnderstandingRegion[] spanning = regions.Where(region => region.XEnd - region.XStart >= pageWidth * 0.62D).OrderByDescending(static region => region.YTop).ToArray();
            var ordered = new List<PdfUnderstandingRegion>(regions.Count);
            var consumed = new HashSet<PdfUnderstandingRegion>();
            double bandTop = double.PositiveInfinity;
            foreach (PdfUnderstandingRegion divider in spanning) {
                AddBand(regions.Where(region => !consumed.Contains(region) && !ReferenceEquals(region, divider) && region.YTop <= bandTop && region.YTop > divider.YTop));
                ordered.Add(divider); consumed.Add(divider); bandTop = divider.YBottom;
            }
            AddBand(regions.Where(region => !consumed.Contains(region)));
            return ordered.ToArray();

            void AddBand(IEnumerable<PdfUnderstandingRegion> candidates) {
                PdfUnderstandingRegion[] band = candidates.ToArray();
                if (band.Length == 0) return;
                double medianWidth = band.Select(region => region.XEnd - region.XStart).OrderBy(static width => width).ElementAt(band.Length / 2);
                double columnTolerance = Math.Max(18D, medianWidth * 0.2D);
                var columns = new List<List<PdfUnderstandingRegion>>();
                foreach (PdfUnderstandingRegion region in band.OrderBy(static region => region.XStart)) {
                    List<PdfUnderstandingRegion>? column = columns.FirstOrDefault(candidate => Math.Abs(candidate.Average(static item => item.XStart) - region.XStart) <= columnTolerance);
                    if (column is null) { column = new List<PdfUnderstandingRegion>(); columns.Add(column); }
                    column.Add(region);
                }
                foreach (List<PdfUnderstandingRegion> column in columns.OrderBy(static column => column.Min(static region => region.XStart))) {
                    foreach (PdfUnderstandingRegion region in column.OrderByDescending(static region => region.YTop).ThenBy(static region => region.XStart)) {
                        if (consumed.Add(region)) ordered.Add(region);
                    }
                }
            }
        }
    }

    private sealed class AdvancedSemanticClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) {
            double[] sizes = orderedRegions.SelectMany(static region => region.Lines).Select(static line => line.FontSize).OrderBy(static size => size).ToArray();
            double median = sizes.Length == 0 ? 0D : sizes[sizes.Length / 2];
            var result = new List<PdfUnderstandingSemanticElement>(orderedRegions.Count);
            foreach (PdfUnderstandingRegion region in orderedRegions) {
                (PdfUnderstandingSemanticKind kind, double confidence, string code, string message) = Classify(context, region, median);
                result.Add(new PdfUnderstandingSemanticElement(region, kind, confidence, new[] { new PdfInferenceEvidence(code, message, confidence - 0.5D) }));
            }
            return result.AsReadOnly();
        }

        private static (PdfUnderstandingSemanticKind Kind, double Confidence, string Code, string Message) Classify(PdfUnderstandingPageContext context, PdfUnderstandingRegion region, double median) {
            string text = region.Text.Trim();
            double largest = region.Lines.Max(static line => line.FontSize);
            if (region.YTop >= context.Height * 0.94D) return (PdfUnderstandingSemanticKind.Header, 0.82D, "semantic.page-edge-header", "The region occupies the top six percent of the page.");
            if (region.YBottom <= context.Height * 0.08D && median > 0D && largest <= median * 0.9D) return (PdfUnderstandingSemanticKind.Footnote, 0.84D, "semantic.bottom-small-text", "Small text occupies the bottom eight percent of the page.");
            if (region.YBottom <= context.Height * 0.05D) return (PdfUnderstandingSemanticKind.Footer, 0.78D, "semantic.page-edge-footer", "The region occupies the bottom five percent of the page.");
            if (text.StartsWith("Figure ", StringComparison.OrdinalIgnoreCase) || text.StartsWith("Fig. ", StringComparison.OrdinalIgnoreCase) || text.StartsWith("Table ", StringComparison.OrdinalIgnoreCase)) return (PdfUnderstandingSemanticKind.Caption, 0.9D, "semantic.caption-prefix", "The region starts with a conventional figure or table caption prefix.");
            if (LooksLikeTable(region)) return (PdfUnderstandingSemanticKind.Table, 0.83D, "semantic.column-alignment", "Several lines share aligned word columns with large horizontal gaps.");
            if (StartsWithListMarker(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, "semantic.list-marker", "The region begins with a bullet or numbered marker.");
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, 0.82D, "semantic.large-font", "The region font is materially larger than the page median.");
            return (PdfUnderstandingSemanticKind.Paragraph, 0.72D, "semantic.body-region", "No stronger business-document semantic signal was found.");
        }

        private static bool LooksLikeTable(PdfUnderstandingRegion region) {
            if (region.Lines.Count < 2) return false;
            int alignedRows = 0;
            foreach (PdfUnderstandingLine line in region.Lines) {
                if (line.Words.Count < 2) continue;
                int largeGaps = 0;
                for (int i = 1; i < line.Words.Count; i++) if (line.Words[i].XStart - line.Words[i - 1].XEnd >= Math.Max(12D, line.FontSize)) largeGaps++;
                if (largeGaps > 0) alignedRows++;
            }
            return alignedRows >= 2;
        }

        private static bool StartsWithListMarker(string text) {
            if (text.Length == 0) return false;
            if (text[0] == '-' || text[0] == '*' || text[0] == '•') return true;
            int index = 0; while (index < text.Length && char.IsDigit(text[index])) index++;
            return index > 0 && index < text.Length && (text[index] == '.' || text[index] == ')');
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
