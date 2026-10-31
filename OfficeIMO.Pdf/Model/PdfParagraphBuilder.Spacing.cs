namespace OfficeIMO.Pdf;

public sealed partial class PdfParagraphBuilder {
    private double _currentHorizontalTextScaling = 100D;
    private global::OfficeIMO.OfficeLength _currentCharacterSpacing;

    private PdfTextRun ApplyCurrentTextSpacing(PdfTextRun run) =>
        run.HorizontalTextScaling == _currentHorizontalTextScaling && run.CharacterSpacing == _currentCharacterSpacing
            ? run : run.WithHorizontalTextScaling(_currentHorizontalTextScaling).WithCharacterSpacing(_currentCharacterSpacing);

    /// <summary>Sets the horizontal glyph width percentage for subsequent text. Use 100 for normal width.</summary>
    public PdfParagraphBuilder HorizontalTextScaling(double percentage) {
        Guard.Positive(percentage, nameof(percentage));
        _currentHorizontalTextScaling = percentage;
        return this;
    }

    /// <summary>Sets additional advance after each rendered glyph, as a physical length, for subsequent text.</summary>
    /// <remarks>The spacing is independent of horizontal glyph scaling. Use zero to restore normal spacing.</remarks>
    public PdfParagraphBuilder CharacterSpacing(global::OfficeIMO.OfficeLength spacing) {
        _currentCharacterSpacing = spacing;
        return this;
    }
}
