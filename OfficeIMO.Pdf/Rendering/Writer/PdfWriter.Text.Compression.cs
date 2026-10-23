using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>One native-source line's immutable admission result; authored and natural advances remain unchanged.</summary>
    private sealed class WordCompressionPlan {
        internal WordCompressionPlan(double naturalWidth, double availableWidth, double[] deltas) {
            NaturalWidth = naturalWidth;
            AvailableWidth = availableWidth;
            LeadingDeltas = Array.AsReadOnly((double[])deltas.Clone());
            Width = naturalWidth + deltas.Sum();
            HasCompression = deltas.Any(delta => delta < 0D);
        }
        internal double NaturalWidth { get; }
        internal double AvailableWidth { get; }
        internal double Width { get; }
        internal bool HasCompression { get; }
        internal System.Collections.ObjectModel.ReadOnlyCollection<double> LeadingDeltas { get; }
    }

    // A tab anchors its field independently of earlier text. Keep the existing
    // anchored-field contract: only ordinary spaces after a final left tab move.
    private static int WordCompressionStart(IReadOnlyList<RichSeg> candidate) {
        for (int index = candidate.Count - 1; index >= 0; index--) {
            if (!candidate[index].LeadingIsTab) continue;
            return candidate[index].LeadingTabAlignment == PdfTabAlignment.Left ? index + 1 : candidate.Count;
        }
        return 0;
    }

    private static double WordCompressionCapacity(RichSeg segment) =>
        segment.LeadingSpace && segment.LeadingIsOrdinarySpace && !segment.LeadingIsTab &&
        segment.LeadingSpaceIsExpandable && segment.LeadingAdvance > 0D
            ? segment.LeadingAdvance / 4D : 0D;

    /// <summary>Measures the unchanged ordinary-space allowance without consuming candidate source state.</summary>
    private static double MeasureWordCompressionCapacity(IReadOnlyList<RichSeg> candidate) {
        double capacity = 0D;
        for (int index = WordCompressionStart(candidate); index < candidate.Count; index++)
            capacity += WordCompressionCapacity(candidate[index]);
        return capacity;
    }

    /// <summary>Tests the complete natural candidate in current PDF points, without allocating or consuming source state.</summary>
    private static double MeasureWordCompressionMinimumWidth(IReadOnlyList<RichSeg> candidate, double naturalWidth) =>
        naturalWidth - MeasureWordCompressionCapacity(candidate);

    /// <summary>Snapshots pending source separators for a trial; footnote ownership is transferred only by actual consumption.</summary>
    private static List<RichSeg> WordCompressionCandidate(IReadOnlyList<RichSeg> line,
        IReadOnlyList<PendingSpaceFragment> pending, double pendingAdvance, int pendingCount,
        bool pendingTab, bool pendingExpandable, bool pendingOrdinarySpace, PdfTabAlignment tabAlignment,
        RichSeg token, IReadOnlyList<RichSeg>? continuation) {
        var candidate = new List<RichSeg>(line.Count + pending.Count + 1 + (continuation?.Count ?? 0));
        candidate.AddRange(line);
        if (pending.Count > 0) {
            foreach (PendingSpaceFragment fragment in pending) {
                RichSeg source = fragment.Style;
                candidate.Add(new RichSeg(string.Empty, false, false, false, false, null, null, null, null, null,
                    source.Font, source.FontSize, source.Baseline, 0D, leadingSpace: true,
                    leadingAdvance: fragment.Advance, leadingSpaceIsExpandable: source.LeadingSpaceIsExpandable,
                    namedFont: source.NamedFont, leadingIsTab: source.LeadingIsTab,
                    leadingTabAlignment: source.LeadingTabAlignment, leadingSpaceCount: fragment.Count,
                    leadingIsOrdinarySpace: source.LeadingIsOrdinarySpace));
            }
        } else if (pendingAdvance != 0D || pendingCount > 0 || pendingTab) {
            candidate.Add(new RichSeg(string.Empty, false, false, false, false, null, null, null, null, null,
                token.Font, token.FontSize, token.Baseline, 0D, leadingSpace: true,
                leadingAdvance: pendingAdvance, leadingSpaceIsExpandable: pendingExpandable,
                leadingIsTab: pendingTab, leadingTabAlignment: tabAlignment, leadingSpaceCount: pendingCount,
                leadingIsOrdinarySpace: pendingOrdinarySpace));
        }
        candidate.Add(token);
        if (continuation != null) candidate.AddRange(continuation);
        return candidate;
    }

    /// <summary>Binds the final kerned natural line to its actual admission frame, including zero-demand native lines.</summary>
    private static void BindWordCompressionPlan(List<RichSeg> line, double availableWidth) {
        double naturalWidth = MeasureRichLineWidth(line);
        var deltas = new double[line.Count];
        if (naturalWidth > availableWidth &&
            MeasureWordCompressionMinimumWidth(line, naturalWidth) <= availableWidth + RichTextWidthTolerance) {
            var resources = new List<OfficeTextSpacingResource>();
            for (int index = WordCompressionStart(line); index < line.Count; index++) {
                double capacity = WordCompressionCapacity(line[index]);
                if (capacity > 0D) resources.Add(new OfficeTextSpacingResource(index, 0, capacity, 0D));
            }
            OfficeTextSpacingAllocation allocation = OfficeTextSpacingAllocator.Allocate(line.Count, resources, availableWidth - naturalWidth);
            for (int index = 0; index < deltas.Length; index++) deltas[index] = allocation.Deltas[index];
        }
        var plan = new WordCompressionPlan(naturalWidth, availableWidth, deltas);
        for (int index = 0; index < line.Count; index++) line[index] = line[index].WithCompressionPlan(plan);
    }

    /// <summary>Resolves an accepted negative plan for paint. A different frame or changed natural line must be rebound by layout.</summary>
    private static bool TryApplyWordCompressionPlan(List<RichSeg> segments, double availableWidth,
        out RichLineJustification result) {
        result = default;
        if (segments.Count == 0 || segments[0].CompressionPlan is not { HasCompression: true } plan ||
            plan.AvailableWidth != availableWidth || plan.LeadingDeltas.Count != segments.Count ||
            plan.NaturalWidth != MeasureRichLineWidth(segments) ||
            segments.Any(segment => !ReferenceEquals(segment.CompressionPlan, plan))) return false;
        var adjusted = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            double delta = plan.LeadingDeltas[index];
            adjusted.Add(delta == 0D ? segment : segment.WithJustification(0D, segment.LeadingAdvance + delta, null));
        }
        result = new RichLineJustification(adjusted, 0D, plan.Width);
        return true;
    }
}
