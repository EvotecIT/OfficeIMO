using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableParagraphFormatting ReadSupportedCellParagraphStyleFormatting(Paragraph paragraph, MainDocumentPart mainPart) {
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            // Unstyled cells retain table defaults ahead of the document's Normal defaults.
            if (string.IsNullOrWhiteSpace(styleId) || string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase)) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            var formatting = LegacyDocWritableParagraphFormatting.Plain;
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            IEnumerable<Style> styles = mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>();
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                Style? style = styles.FirstOrDefault(candidate => candidate.Type?.Value == StyleValues.Paragraph &&
                    string.Equals(candidate.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
                if (style == null) break;
                LegacyDocWritableParagraphFormatting own = TryMapBuiltInParagraphStyleIndex(styleId!, out ushort styleIndex)
                    ? ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.StyleParagraphProperties)
                    : ReadSupportedCustomParagraphStyleParagraphFormatting(style.StyleParagraphProperties);
                formatting = formatting.WithInheritedParagraphFormatting(own);
                styleId = style.BasedOn?.Val?.Value;
            }
            return formatting;
        }
    }
}
