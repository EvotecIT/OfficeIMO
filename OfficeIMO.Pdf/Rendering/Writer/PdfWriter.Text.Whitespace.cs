namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Coalesce adjacent spaces from the same source run. Retaining run identity
    // avoids per-character objects and keeps decorated/mixed-font continuations stable.
    private sealed class PendingSpaceFragment {
        internal PendingSpaceFragment(RichSeg style) { Style = style; AddSpace(); }
        internal RichSeg Style { get; }
        internal double Advance { get; set; }
        internal int Count { get; private set; }
        internal void AddSpace() { Count++; Advance += Style.LeadingAdvance; }
    }

    private static List<RichSeg> ConsumePreservedSpaceFragments(List<PendingSpaceFragment> pending, double advance) {
        var result = new List<RichSeg>();
        double remaining = advance;
        while (pending.Count > 0) {
            PendingSpaceFragment fragment = pending[0];
            RichSeg source = fragment.Style;
            if (remaining <= 0.001D && fragment.Advance > 0.001D) break;
            double consumed = fragment.Advance > 0D ? Math.Min(fragment.Advance, remaining) : fragment.Advance;
            double unit = source.LeadingAdvance;
            int count = unit > 0D ? (int)Math.Floor(consumed / unit + 0.00000001D) : fragment.Count;
            double remainder = unit > 0D ? Math.Max(0D, consumed - count * unit) : 0D;
            result.Add(new RichSeg(string.Empty, source.Bold, source.Italic, source.Underline, source.Strike,
                source.Color, source.BackgroundColor, source.Uri, source.DestinationName, source.Contents,
                source.Font, source.FontSize, source.Baseline, 0D, leadingSpace: true,
                leadingAdvance: consumed, leadingSpaceIsExpandable: source.LeadingSpaceIsExpandable,
                namedFont: source.NamedFont, underlineStyle: source.UnderlineStyle, strikeStyle: source.StrikeStyle,
                decorationColor: source.DecorationColor, featureSettings: source.FeatureSettings,
                textDirection: source.TextDirection, leadingUnderlineStyle: source.LeadingUnderlineStyle,
                leadingDecorationColor: source.LeadingDecorationColor, leadingDecorationFontSize: source.LeadingDecorationFontSize,
                leadingDecorationTextRise: source.LeadingDecorationTextRise,
                horizontalTextScaling: source.HorizontalTextScaling, characterSpacing: source.CharacterSpacing,
                fontMetricScale: source.FontMetricScale, leadingSpaceCount: count,
                leadingSpaceRemainder: remainder, leadingSpaceUnitAdvance: unit));
            remaining -= consumed;
            fragment.Advance -= consumed;
            if (fragment.Advance <= 0.001D) pending.RemoveAt(0);
            else break;
        }
        return result;
    }

    private static void AppendPreservedSpaceRuns(List<PdfTextRun> runs, RichSeg segment) {
        // Links remain on their visible text. A whitespace-only public run
        // cannot carry a link target, and it does not need a separate annotation.
        RichSeg spaceStyle = segment.WithoutLink();
        if (segment.LeadingSpaceCount > 0)
            runs.Add(BuildTextRunFromWrappedSegment(new string(' ', segment.LeadingSpaceCount), spaceStyle));
        if (segment.LeadingSpaceRemainder > 0.001D) {
            // A literal space can straddle a frame edge. Its remainder has the
            // same font/decorations and only the advance that this fragment owns.
            runs.Add(BuildTextRunFromWrappedSegment(" ", spaceStyle).WithCharacterSpacing(
                segment.CharacterSpacing + segment.LeadingSpaceRemainder - segment.LeadingSpaceUnitAdvance));
        }
    }
}
