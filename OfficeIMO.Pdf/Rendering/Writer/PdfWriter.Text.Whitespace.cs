namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Coalesce adjacent spaces from the same source run. Retaining run identity
    // avoids per-character objects and keeps decorated/mixed-font continuations stable.
    private sealed class PendingSpaceFragment {
        internal PendingSpaceFragment(RichSeg style) {
            Style = style;
            if (style.LeadingIsTab) Advance = style.LeadingAdvance;
            else AddSpace();
        }
        internal RichSeg Style { get; }
        internal double Advance { get; set; }
        internal int Count { get; private set; }
        internal void AddSpace() { Count++; Advance += Style.LeadingAdvance; }
    }

    private static List<RichSeg> ConsumePreservedSpaceFragments(List<PendingSpaceFragment> pending, double advance) {
        var result = new List<RichSeg>();
        // A complete group has a signed net advance. A positive prefix can
        // exceed that net when a later source run has negative tracking.
        bool consumeAll = Math.Abs(advance - pending.Sum(fragment => fragment.Advance)) <= 0.001D;
        double remaining = advance;
        while (pending.Count > 0) {
            PendingSpaceFragment fragment = pending[0];
            RichSeg source = fragment.Style;
            if (!consumeAll && remaining <= 0.001D && fragment.Advance > 0.001D) break;
            double consumed = !consumeAll && fragment.Advance > 0D ? Math.Min(fragment.Advance, remaining) : fragment.Advance;
            double unit = source.LeadingAdvance;
            int count = source.LeadingIsTab ? 0 : unit > 0D ? (int)Math.Floor(consumed / unit + 0.00000001D) : fragment.Count;
            double remainder = !source.LeadingIsTab && unit > 0D ? Math.Max(0D, consumed - count * unit) : 0D;
            result.Add(new RichSeg(string.Empty, source.Bold, source.Italic, source.Underline, source.Strike,
                source.Color, source.BackgroundColor, source.Uri, source.DestinationName, source.Contents,
                source.Font, source.FontSize, source.Baseline, 0D, leadingSpace: true,
                leadingAdvance: consumed, leadingSpaceIsExpandable: source.LeadingSpaceIsExpandable,
                leadingTabLeader: source.LeadingTabLeader, leadingTabStop: source.LeadingTabStop,
                leadingIsTab: source.LeadingIsTab, leadingTabAlignment: source.LeadingTabAlignment,
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

    private static RichSeg CreatePreservedTabFragment(RichSeg style, double advance,
        PdfTabStop? stop, PdfTabAlignment alignment, PdfTabLeaderStyle leader) =>
        new RichSeg(string.Empty, style.Bold, style.Italic, style.Underline, style.Strike,
            style.Color, style.BackgroundColor, style.Uri, style.DestinationName, style.Contents,
            style.Font, style.FontSize, style.Baseline, 0D, leadingSpace: true,
            leadingAdvance: advance, leadingSpaceIsExpandable: false, leadingTabLeader: leader,
            namedFont: style.NamedFont, underlineStyle: style.UnderlineStyle, strikeStyle: style.StrikeStyle,
            decorationColor: style.DecorationColor, featureSettings: style.FeatureSettings,
            textDirection: style.TextDirection, leadingTabStop: stop, leadingIsTab: true,
            leadingTabAlignment: alignment, horizontalTextScaling: style.HorizontalTextScaling,
            characterSpacing: style.CharacterSpacing, fontMetricScale: style.FontMetricScale,
            leadingSpaceCount: 0);

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
