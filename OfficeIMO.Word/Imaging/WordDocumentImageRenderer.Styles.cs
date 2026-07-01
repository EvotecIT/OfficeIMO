using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static IEnumerable<StyleParagraphProperties> EnumerateParagraphStyleProperties(WordParagraph paragraph) {
            foreach (Style style in EnumerateParagraphStyleChain(paragraph)) {
                if (style.StyleParagraphProperties != null) {
                    yield return style.StyleParagraphProperties;
                }
            }
        }

        private static IEnumerable<StyleRunProperties> EnumerateRunStyleProperties(WordParagraph paragraph) {
            foreach (Style style in EnumerateCharacterStyleChain(paragraph)) {
                if (style.StyleRunProperties != null) {
                    yield return style.StyleRunProperties;
                }
            }

            foreach (Style style in EnumerateParagraphStyleChain(paragraph)) {
                if (style.StyleRunProperties != null) {
                    yield return style.StyleRunProperties;
                }
            }
        }

        private static IEnumerable<StyleTableProperties> EnumerateTableStyleProperties(WordTable table) {
            foreach (Style style in EnumerateTableStyleChain(table)) {
                if (style.StyleTableProperties != null) {
                    yield return style.StyleTableProperties;
                }
            }
        }

        private static IEnumerable<TableStyleProperties> EnumerateTableConditionalStyleProperties(WordTable table, TableStyleOverrideValues type) {
            foreach (Style style in EnumerateTableStyleChain(table)) {
                foreach (TableStyleProperties properties in style.Elements<TableStyleProperties>()) {
                    if (properties.Type?.Value == type) {
                        yield return properties;
                    }
                }
            }
        }

        private static IEnumerable<Style> EnumerateParagraphStyleChain(WordParagraph paragraph) =>
            EnumerateStyleChain(GetDocumentStyles(paragraph), paragraph.StyleId, StyleValues.Paragraph);

        private static IEnumerable<Style> EnumerateCharacterStyleChain(WordParagraph paragraph) =>
            EnumerateStyleChain(GetDocumentStyles(paragraph), paragraph.CharacterStyleId, StyleValues.Character);

        private static IEnumerable<Style> EnumerateTableStyleChain(WordTable table) =>
            EnumerateStyleChain(GetDocumentStyles(table), table._tableProperties?.TableStyle?.Val?.Value, StyleValues.Table);

        private static Styles? GetDocumentStyles(WordParagraph paragraph) =>
            paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;

        private static Styles? GetDocumentStyles(WordTable table) =>
            table.Document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;

        private static IEnumerable<Style> EnumerateStyleChain(Styles? styles, string? styleId, StyleValues expectedType) {
            if (styles == null || string.IsNullOrWhiteSpace(styleId)) {
                yield break;
            }

            HashSet<string> visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                Style? style = styles.Elements<Style>()
                    .FirstOrDefault(candidate =>
                        string.Equals(candidate.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase) &&
                        (candidate.Type == null || !candidate.Type.HasValue || candidate.Type.Value == expectedType));
                if (style == null) {
                    yield break;
                }

                yield return style;
                styleId = style.BasedOn?.Val?.Value;
            }
        }
    }
}
