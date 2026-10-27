namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    internal PdfEmbeddedFontFallbackCandidate? PlannedFontFallback { get; private set; }

    internal PdfTextRun WithPlannedFontFallback(PdfEmbeddedFontFallbackCandidate? candidate) {
        PlannedFontFallback = candidate;
        return this;
    }

    // Resolve the plan before wrapping so the ordinary named-font path measures
    // and paints the same resource, while preserving all other run metadata.
    internal PdfTextRun WithResolvedFallbackFontFamily(string familyName) {
        familyName = OfficeIMO.Drawing.OfficeFontFamilyParser.SerializeLiteral(familyName);
        PdfTextRun copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic,
            Strike, FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName,
            TabLeader, TabAlignment, BackgroundColor, familyName, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.PlannedFontFallback = null;
        return copy;
    }
}
