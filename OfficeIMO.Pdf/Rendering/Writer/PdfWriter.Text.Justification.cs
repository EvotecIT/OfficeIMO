using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Source paragraph identity carried through cell line slicing and physical fragments.</summary>
    private readonly record struct RichLineJustificationPolicy(PdfJustificationMode Mode,
        bool ExplicitBreaks, bool LastLine, bool IsFinalParagraphLine, bool IsExplicitBreak,
        bool UseWordJustificationSpacing = false) {
        internal RichLineJustificationPolicy(PdfParagraphStyle? style, bool final, bool explicitBreak)
            : this(style?.JustificationMode ?? PdfJustificationMode.Words,
                style?.JustifyExplicitLineBreaks == true, style?.JustifyLastLine == true, final, explicitBreak,
                style?.UseWordJustificationSpacing == true) { }
    }

    private readonly struct RichLineJustification {
        internal RichLineJustification(List<RichSeg> segments, double wordSpacing, double width) {
            Segments = segments;
            WordSpacing = wordSpacing;
            Width = width;
        }
        internal List<RichSeg> Segments { get; }
        internal double WordSpacing { get; }
        internal double Width { get; }
    }

    /// <summary>Plans one physical line's glyph and separator advances for paint and paint diagnostics.</summary>
    private static RichLineJustification PlanRichLineJustification(List<RichSeg> segments,
        PdfAlign alignment, PdfParagraphStyle? style, double availableWidth, bool isFinalParagraphLine, PdfOptions options) {
        if (alignment == PdfAlign.Justify &&
            TryApplyWordCompressionPlan(segments, availableWidth, out RichLineJustification compression))
            return compression;
        double naturalWidth = MeasureRichLineWidth(segments);
        bool hardBreak = segments.Any(segment => segment.EndsWithHardBreak);
        RichLineJustificationPolicy policy = segments.Select(segment => segment.JustificationPolicy).FirstOrDefault(item => item.HasValue)
            ?? new RichLineJustificationPolicy(style, isFinalParagraphLine, hardBreak);
        bool expands = alignment == PdfAlign.Justify && availableWidth > naturalWidth &&
            (policy.IsExplicitBreak ? policy.ExplicitBreaks : !policy.IsFinalParagraphLine || policy.LastLine);
        if (!expands) return new RichLineJustification(segments, 0D, naturalWidth);
        int lastTab = segments.FindLastIndex(segment => segment.LeadingIsTab);
        if (lastTab >= 0) {
            if (segments[lastTab].LeadingTabAlignment != PdfTabAlignment.Left)
                return new RichLineJustification(segments, 0D, naturalWidth);
            if (policy.UseWordJustificationSpacing && policy.Mode == PdfJustificationMode.Words)
                return PlanRichWordSpaceDistribution(segments, naturalWidth, availableWidth, lastTab + 1);
            return PlanRichAnchoredWordDistribution(segments, lastTab, naturalWidth, availableWidth);
        }
        if (policy.Mode != PdfJustificationMode.Characters) {
            if (policy.UseWordJustificationSpacing)
                return PlanRichWordSpaceDistribution(segments, naturalWidth, availableWidth);
            int gaps = segments.Where(segment => segment.LeadingSpace && segment.LeadingSpaceIsExpandable)
                .Sum(segment => segment.LeadingSpaceCount);
            return gaps > 0 ? new RichLineJustification(segments, (availableWidth - naturalWidth) / gaps, availableWidth)
                : new RichLineJustification(segments, 0D, naturalWidth);
        }
        return PlanRichCharacterDistribution(segments, naturalWidth, availableWidth, options, policy.UseWordJustificationSpacing);
    }

    /// <summary>Consumes source-space preferred expansion before sharing the remaining width per ordinary space.</summary>
    private static RichLineJustification PlanRichWordSpaceDistribution(List<RichSeg> segments,
        double naturalWidth, double availableWidth, int startIndex = 0) {
        int spaces = segments.Skip(startIndex).Where(IsWordExpansionSpace).Sum(segment => segment.LeadingSpaceCount);
        if (spaces == 0) return new RichLineJustification(segments, 0D, naturalWidth);
        double adjustment = availableWidth - naturalWidth;
        var resources = new List<OfficeTextSpacingResource>();
        for (int index = startIndex; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            if (!IsWordExpansionSpace(segment)) continue;
            // Preferred capacity follows the actual source-space metrics. The later
            // tier needs only the finite remaining request, shared per source space;
            // it does not introduce a font-specific or native-unit ceiling.
            resources.Add(new OfficeTextSpacingResource(index, 0, 0D, WordPreferredExpansionCapacity(segment)));
            resources.Add(new OfficeTextSpacingResource(index, 4, 0D, adjustment * (segment.LeadingSpaceCount / (double)spaces)));
        }
        OfficeTextSpacingAllocation allocation = OfficeTextSpacingAllocator.Allocate(segments.Count, resources, adjustment);
        var distributed = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            double delta = allocation.Deltas[index];
            distributed.Add(delta == 0D ? segment : segment.WithJustification(0D, segment.LeadingAdvance + delta, null));
        }
        return new RichLineJustification(distributed, 0D, naturalWidth + allocation.AppliedAdjustment);
    }

    private static bool IsWordExpansionSpace(RichSeg segment) => segment.LeadingSpace &&
        segment.LeadingIsOrdinarySpace && !segment.LeadingIsTab && segment.LeadingSpaceIsExpandable && segment.LeadingSpaceCount > 0;

    /// <summary>The source-space preferred tier shared by positive paint and whole-word break selection.</summary>
    private static double WordPreferredExpansionCapacity(RichSeg segment) =>
        IsWordExpansionSpace(segment) ? Math.Max(0D, segment.LeadingAdvance / 2D) : 0D;

    private static RichLineJustification PlanRichAnchoredWordDistribution(List<RichSeg> segments,
        int lastTab, double naturalWidth, double availableWidth) {
        int gaps = 0;
        for (int index = lastTab + 1; index < segments.Count; index++)
            if (segments[index].LeadingSpace && segments[index].LeadingSpaceIsExpandable)
                gaps += segments[index].LeadingSpaceCount;
        if (gaps == 0) return new RichLineJustification(segments, 0D, naturalWidth);
        double spacing = (availableWidth - naturalWidth) / gaps;
        var distributed = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            // A left tab fixes the following field's start. Both policies retain all earlier
            // advances and expand only word separators within the final left-tab field.
            bool expands = index > lastTab && segment.LeadingSpace && segment.LeadingSpaceIsExpandable;
            distributed.Add(expands ? segment.WithJustification(0D,
                segment.LeadingAdvance + segment.LeadingSpaceCount * spacing, null) : segment);
        }
        return new RichLineJustification(distributed, 0D, availableWidth);
    }

    /// <summary>Maps logical ownership back to the encoded visual order, including RTL glyphs.</summary>
    private static int[] ResolveRichGlyphClusterStarts(IReadOnlyList<PdfGlyphInfo> glyphs) {
        int[] order = Enumerable.Range(0, glyphs.Count)
            .OrderBy(index => glyphs[index].LogicalClusterStart).ThenBy(index => glyphs[index].TextIndex).ToArray();
        PdfGlyphInfo[] logical = order.Select(index => glyphs[index]).ToArray();
        var starts = new int[glyphs.Count];
        for (int start = 0; start < logical.Length;) {
            int end = PdfGlyphClusters.End(logical, start, logical.Length);
            int sourceIndex = logical[start].LogicalClusterStart;
            for (int index = start; index < end; index++) starts[order[index]] = sourceIndex;
            start = end;
        }
        return starts;
    }
}
