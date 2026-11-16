using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>
        /// Resolves Word's automatic text color against the solid fill the exporter paints behind the text: a body
        /// paragraph's shading panel or the containing table cell. Returns white on a dark fill and
        /// <see langword="null"/> (the ordinary black default) otherwise.
        /// </summary>
        private static global::OfficeIMO.OfficeColor? ResolveNativeAutomaticTextColor(WordParagraph paragraph) {
            global::OfficeIMO.OfficeColor? background = PaintsNativeParagraphShadingPanel(paragraph)
                ? ParseNativeColor(GetNativeEffectiveParagraphShadingFill(paragraph))
                : null;
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

        // Only ordinary body paragraphs render their shading as a filled panel. Headings, table-cell
        // paragraphs, text boxes and running content draw the text without the paragraph fill.
        private static bool PaintsNativeParagraphShadingPanel(WordParagraph paragraph) {
            W.Paragraph? source = paragraph._paragraph;
            if (source == null || GetHeadingLevel(paragraph) > 0) return false;
            foreach (var ancestor in source.Ancestors()) {
                if (ancestor is W.TableCell or W.TextBoxContent or W.Header or W.Footer or W.Footnote or W.Endnote) return false;
                if (ancestor is W.Body) return true;
            }
            return false;
        }
    }
}
