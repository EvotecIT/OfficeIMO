namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool HasRichTextSpacing(double horizontalTextScaling, double characterSpacing) =>
        horizontalTextScaling != 100D || characterSpacing != 0D;

    private static void ApplyRichTextSpacing(ContentStreamBuilder content, double horizontalTextScaling, double characterSpacing, double wordSpacing = 0D) {
        if (!HasRichTextSpacing(horizontalTextScaling, characterSpacing)) return;
        double scale = horizontalTextScaling / 100D;
        // PDF scales Tc and Tw with glyph widths. The public run contract uses page-point advances.
        content.HorizontalTextScaling(horizontalTextScaling)
            .CharacterSpacing(characterSpacing / scale)
            .WordSpacing(wordSpacing / scale);
    }

    private static void ResetRichTextSpacing(ContentStreamBuilder content, double horizontalTextScaling, double characterSpacing, double wordSpacing = 0D) {
        if (!HasRichTextSpacing(horizontalTextScaling, characterSpacing)) return;
        content.HorizontalTextScaling(100D).CharacterSpacing(0D).WordSpacing(wordSpacing);
    }
}
