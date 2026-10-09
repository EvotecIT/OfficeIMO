namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Horizontal glyph width as a percentage of the selected font's normal width. The default is 100.</summary>
    public double HorizontalTextScaling { get; private set; } = 100D;

    /// <summary>Additional advance after each rendered glyph, in page points. The default is zero.</summary>
    /// <remarks>This advance is independent of <see cref="HorizontalTextScaling"/> and does not change glyph height.</remarks>
    public double CharacterSpacing { get; private set; }

    /// <summary>Creates a copy with a positive horizontal glyph width percentage, preserving font height and other formatting.</summary>
    public PdfTextRun WithHorizontalTextScaling(double percentage) {
        Guard.Positive(percentage, nameof(percentage));
        if (InlineElement != null) return this;
        PdfTextRun copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic, Strike, FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader, TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.HorizontalTextScaling = percentage;
        return copy;
    }

    /// <summary>Creates a copy with additional advance after each rendered glyph, in page points.</summary>
    /// <param name="points">Finite positive or negative spacing. Negative values condense the glyph advances.</param>
    public PdfTextRun WithCharacterSpacing(double points) {
        if (double.IsNaN(points) || double.IsInfinity(points)) {
            throw new System.ArgumentOutOfRangeException(nameof(points), "Character spacing must be finite.");
        }
        if (InlineElement != null) return this;
        PdfTextRun copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic, Strike, FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader, TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.CharacterSpacing = points;
        return copy;
    }

    internal PdfTextRun WithSpacingFrom(PdfTextRun source) {
        HorizontalTextScaling = source.HorizontalTextScaling;
        CharacterSpacing = source.CharacterSpacing;
        return this;
    }

    internal PdfTextRun WithInheritedFontStyle(bool bold, bool italic) {
        if (InlineElement != null || ((!bold || Bold) && (!italic || Italic))) return this;
        return CopyRunFormattingTo(new PdfTextRun(Text, Bold || bold, Underline, Color, Italic || italic,
            Strike, FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader,
            TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
    }

    private PdfTextRun CopyRunFormattingTo(PdfTextRun copy) {
        copy.FeatureSettings = FeatureSettings;
        copy.HorizontalOffset = HorizontalOffset;
        copy.TextDirection = TextDirection;
        return copy.WithSpacingFrom(this);
    }
}
