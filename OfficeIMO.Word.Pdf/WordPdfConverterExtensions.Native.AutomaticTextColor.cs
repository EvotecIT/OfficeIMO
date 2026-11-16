using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>
        /// Resolves Word's automatic text color against a solid paragraph or table-cell fill. Returns white on a
        /// dark fill and <see langword="null"/> (the ordinary black default) otherwise.
        /// </summary>
        private static global::OfficeIMO.OfficeColor? ResolveNativeAutomaticTextColor(WordParagraph paragraph) {
            global::OfficeIMO.OfficeColor? background = ParseNativeColor(GetNativeEffectiveParagraphShadingFill(paragraph));
            if (!background.HasValue && paragraph._document != null) {
                W.Shading? cellShading = WordRunFormatting.ResolveContainingTableCellShading(paragraph);
                background = cellShading == null ? null : ParseNativeColor(WordRunFormatting.ResolveShadingFillColorHex(cellShading, paragraph._document));
            }

            if (background is not { } fill) {
                return null;
            }

            // Word switches automatic text on encoded RGB brightness, not WCAG contrast; Word-produced
            // DOC and DOCX references cover both sides of this threshold.
            var rgb = fill.ToNormalizedRgb();
            double brightness = 0.299D * rgb.R + 0.587D * rgb.G + 0.114D * rgb.B;
            return brightness < 60D / 256D ? global::OfficeIMO.OfficeColor.White : null;
        }
    }
}
