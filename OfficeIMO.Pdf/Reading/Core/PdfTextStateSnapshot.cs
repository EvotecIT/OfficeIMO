namespace OfficeIMO.Pdf;

internal readonly struct PdfTextStateSnapshot {
    internal PdfTextStateSnapshot(
        string fontResource,
        double fontSize,
        double leading,
        double characterSpacing,
        double wordSpacing,
        double horizontalScaling,
        double textRise,
        int textRenderingMode) {
        FontResource = fontResource;
        FontSize = fontSize;
        Leading = leading;
        CharacterSpacing = characterSpacing;
        WordSpacing = wordSpacing;
        HorizontalScaling = horizontalScaling;
        TextRise = textRise;
        TextRenderingMode = textRenderingMode;
    }

    internal static PdfTextStateSnapshot Default { get; } = new PdfTextStateSnapshot(
        "F1",
        12D,
        14.4D,
        0D,
        0D,
        1D,
        0D,
        0);

    internal string FontResource { get; }
    internal double FontSize { get; }
    internal double Leading { get; }
    internal double CharacterSpacing { get; }
    internal double WordSpacing { get; }
    internal double HorizontalScaling { get; }
    internal double TextRise { get; }
    internal int TextRenderingMode { get; }

    internal PdfTextStateSnapshot WithFont(string fontResource, double fontSize) =>
        new PdfTextStateSnapshot(fontResource, fontSize, Leading, CharacterSpacing, WordSpacing, HorizontalScaling, TextRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithLeading(double leading) =>
        new PdfTextStateSnapshot(FontResource, FontSize, leading, CharacterSpacing, WordSpacing, HorizontalScaling, TextRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithCharacterSpacing(double characterSpacing) =>
        new PdfTextStateSnapshot(FontResource, FontSize, Leading, characterSpacing, WordSpacing, HorizontalScaling, TextRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithWordSpacing(double wordSpacing) =>
        new PdfTextStateSnapshot(FontResource, FontSize, Leading, CharacterSpacing, wordSpacing, HorizontalScaling, TextRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithHorizontalScaling(double horizontalScaling) =>
        new PdfTextStateSnapshot(FontResource, FontSize, Leading, CharacterSpacing, WordSpacing, horizontalScaling, TextRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithTextRise(double textRise) =>
        new PdfTextStateSnapshot(FontResource, FontSize, Leading, CharacterSpacing, WordSpacing, HorizontalScaling, textRise, TextRenderingMode);

    internal PdfTextStateSnapshot WithTextRenderingMode(int textRenderingMode) =>
        new PdfTextStateSnapshot(FontResource, FontSize, Leading, CharacterSpacing, WordSpacing, HorizontalScaling, TextRise, textRenderingMode);
}
