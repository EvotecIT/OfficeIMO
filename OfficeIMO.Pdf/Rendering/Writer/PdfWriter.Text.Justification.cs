using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Source paragraph identity carried through cell line slicing and physical fragments.</summary>
    private readonly record struct RichLineJustificationPolicy(PdfJustificationMode Mode,
        bool ExplicitBreaks, bool LastLine, bool IsFinalParagraphLine, bool IsExplicitBreak) {
        internal RichLineJustificationPolicy(PdfParagraphStyle? style, bool final, bool explicitBreak)
            : this(style?.JustificationMode ?? PdfJustificationMode.Words,
                style?.JustifyExplicitLineBreaks == true, style?.JustifyLastLine == true, final, explicitBreak) { }
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
        double naturalWidth = MeasureRichLineWidth(segments);
        bool hardBreak = segments.Any(segment => segment.EndsWithHardBreak);
        RichLineJustificationPolicy policy = segments.Select(segment => segment.JustificationPolicy).FirstOrDefault(item => item.HasValue)
            ?? new RichLineJustificationPolicy(style, isFinalParagraphLine, hardBreak);
        bool expands = alignment == PdfAlign.Justify && availableWidth > naturalWidth &&
            (policy.IsExplicitBreak ? policy.ExplicitBreaks : !policy.IsFinalParagraphLine || policy.LastLine);
        if (!expands) return new RichLineJustification(segments, 0D, naturalWidth);
        int lastTab = segments.FindLastIndex(segment => segment.LeadingIsTab);
        if (lastTab >= 0) {
            return segments[lastTab].LeadingTabAlignment == PdfTabAlignment.Left
                ? PlanRichAnchoredWordDistribution(segments, lastTab, naturalWidth, availableWidth)
                : new RichLineJustification(segments, 0D, naturalWidth);
        }
        if (policy.Mode != PdfJustificationMode.Characters) {
            int gaps = segments.Where(segment => segment.LeadingSpace && segment.LeadingSpaceIsExpandable)
                .Sum(segment => segment.LeadingSpaceCount);
            return gaps > 0 ? new RichLineJustification(segments, (availableWidth - naturalWidth) / gaps, availableWidth)
                : new RichLineJustification(segments, 0D, naturalWidth);
        }
        return PlanRichCharacterDistribution(segments, naturalWidth, availableWidth, options);
    }

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

    private static RichLineJustification PlanRichCharacterDistribution(List<RichSeg> segments,
        double naturalWidth, double availableWidth, PdfOptions options) {
        var text = new StringBuilder();
        var indexes = new List<int>();
        var owners = new List<(int Segment, int Glyph)>();
        var commands = new PdfTextShowCommand?[segments.Count];
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            if (segment.LeadingSpace && !segment.LeadingIsTab) {
                for (int space = 0; space < segment.LeadingSpaceCount; space++) {
                    indexes.Add(text.Length);
                    owners.Add((index, -1));
                    text.Append(' ');
                }
            }
            if (segment.InlineElement != null) {
                indexes.Add(text.Length);
                owners.Add((index, -2));
                text.Append('\uFFFC');
                continue;
            }
            if (segment.Text.Length == 0) continue;
            PdfTextShowCommand ownership = segment.KerningCommand is { VisualGlyphs: not null } prepared ? prepared :
                EncodeTextShowCommand(segment.Text, segment.Font, segment.NamedFont, options,
                    segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale,
                    segment.HorizontalTextScaling, segment.PreservesPaintedGlyphs, requireGlyphClusters: true);
            PdfTextShowCommand command = segment.KerningCommand ?? ownership;
            if (command.GlyphCount != ownership.GlyphCount || command.GlyphHex != ownership.GlyphHex)
                throw new InvalidOperationException("Character distribution must retain the prepared encoded glyphs.");
            commands[index] = command;
            int start = text.Length;
            // Encoded glyph ownership is relative to painted text, without optional breaks.
            string paintedText = WithoutOptionalBreaks(segment.Text);
            text.Append(paintedText);
            if (ownership.VisualGlyphs is { } glyphs) {
                int[] sourceIndexes = ResolveRichGlyphClusterStarts(glyphs);
                for (int glyph = 0; glyph < glyphs.Count; glyph++) {
                    indexes.Add(start + sourceIndexes[glyph]);
                    owners.Add((index, glyph));
                }
            } else {
                // The only cluster-free encoding here is a standard one-byte WinAnsi font.
                // Embedded encoders were explicitly requested to return shaped ownership.
                if (command.GlyphCount != paintedText.Length)
                    throw new InvalidOperationException("Character distribution requires encoded glyph source ownership.");
                for (int glyph = 0; glyph < command.GlyphCount; glyph++) {
                    indexes.Add(start + glyph);
                    owners.Add((index, glyph));
                }
            }
        }
        if (indexes.Count < 2) return new RichLineJustification(segments, 0D, naturalWidth);
        bool[] boundaries = OfficeOpenTypeTracking.GetBoundaries(text.ToString(), indexes);
        boundaries[boundaries.Length - 1] = false;
        int opportunities = boundaries.Count(boundary => boundary);
        if (opportunities == 0) return new RichLineJustification(segments, 0D, naturalWidth);
        double spacing = (availableWidth - naturalWidth) / opportunities;
        var leading = new int[segments.Count];
        var inline = new int[segments.Count];
        var glyphBoundaries = new bool[segments.Count][];
        for (int boundary = 0; boundary < boundaries.Length; boundary++) {
            if (!boundaries[boundary]) continue;
            (int segment, int glyph) = owners[boundary];
            if (glyph == -1) leading[segment]++;
            else if (glyph == -2) inline[segment]++;
            else {
                glyphBoundaries[segment] ??= new bool[commands[segment]!.GlyphCount];
                glyphBoundaries[segment][glyph] = true;
            }
        }
        var distributed = new List<RichSeg>(segments.Count);
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            bool[]? selected = glyphBoundaries[index];
            int count = (selected?.Count(boundary => boundary) ?? 0) + inline[index];
            distributed.Add(count == 0 && leading[index] == 0 ? segment : segment.WithJustification(
                count * spacing, segment.LeadingAdvance + leading[index] * spacing,
                selected == null ? null : commands[index]!.WithJustificationTracking(spacing, selected)));
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
