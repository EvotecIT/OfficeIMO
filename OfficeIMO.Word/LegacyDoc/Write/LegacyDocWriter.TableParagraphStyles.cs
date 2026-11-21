using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableParagraphFormatting ReadSupportedCellParagraphStyleFormatting(Paragraph paragraph, MainDocumentPart mainPart) {
            var formatting = LegacyDocWritableParagraphFormatting.Plain;
            foreach (Style style in EnumerateCellParagraphStyleChain(paragraph, mainPart)) {
                string styleId = style.StyleId!.Value!;
                LegacyDocWritableParagraphFormatting own = TryMapBuiltInParagraphStyleIndex(styleId, out ushort styleIndex)
                    ? ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.StyleParagraphProperties)
                    : ReadSupportedCustomParagraphStyleParagraphFormatting(style.StyleParagraphProperties);
                formatting = formatting.WithInheritedParagraphFormatting(own);
            }
            return formatting;
        }

        private static LegacyDocWritableFormatting ReadSupportedCellParagraphStyleRunFormatting(Paragraph paragraph, MainDocumentPart mainPart) {
            var formatting = LegacyDocWritableFormatting.Plain;
            foreach (Style style in EnumerateCellParagraphStyleChain(paragraph, mainPart)) {
                formatting = formatting.WithInheritedFormatting(ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(style)));
            }
            return formatting;
        }

        private static IEnumerable<Style> EnumerateCellParagraphStyleChain(Paragraph paragraph, MainDocumentPart mainPart) {
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            // Unstyled cells retain table defaults ahead of the document's Normal defaults.
            if (string.IsNullOrWhiteSpace(styleId) || string.Equals(styleId, "Normal", StringComparison.Ordinal)) {
                yield break;
            }

            var visited = new HashSet<string>(StringComparer.Ordinal);
            IReadOnlyList<WordStyleDefinitionsReader.StyleDefinition> styles = WordStyleCatalog.For(mainPart).Definitions;
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                WordStyleDefinitionsReader.StyleDefinition? style = styles.FirstOrDefault(candidate => candidate.IsParagraphStyle &&
                    string.Equals(candidate.StyleId, styleId, StringComparison.Ordinal));
                if (style == null) break;
                yield return style.Element;
                styleId = style.BasedOn;
            }
        }
    }
}
