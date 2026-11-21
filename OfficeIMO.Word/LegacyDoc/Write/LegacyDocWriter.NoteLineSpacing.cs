using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static ParagraphProperties? MaterializeNoteLineSpacing(Paragraph paragraph, string noteKind) {
            ParagraphProperties? properties = paragraph.ParagraphProperties;
            if (!string.IsNullOrWhiteSpace(properties?.SpacingBetweenLines?.Line?.Value)) return properties;
            OpenXmlPartRootElement? root = paragraph.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
            MainDocumentPart? mainPart = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
            if (mainPart?.StyleDefinitionsPart == null) return properties;
            WordStyleCatalog definitions = WordStyleCatalog.For(mainPart);
            string? styleId = properties?.ParagraphStyleId?.Val?.Value ?? (noteKind == "footnote" ? "FootnoteText" : "EndnoteText");
            var visited = new HashSet<string>(StringComparer.Ordinal);
            SpacingBetweenLines? inherited = null;
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                if (visited.Count > MaxInheritedStyleDepth) throw new NotSupportedException("Native DOC saving does not support note paragraph style chains beyond 256 levels.");
                Style? style = definitions.FindById(styleId);
                inherited = style?.StyleParagraphProperties?.SpacingBetweenLines;
                if (!string.IsNullOrWhiteSpace(inherited?.Line?.Value)) break;
                styleId = style?.BasedOn?.Val?.Value;
            }
            if (string.IsNullOrWhiteSpace(inherited?.Line?.Value)) inherited = definitions.DocumentDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
            if (string.IsNullOrWhiteSpace(inherited?.Line?.Value)) return properties;
            // DOC uses fixed built-in note style slots. Keep the source's effective
            // spacing on the note paragraph when that style cannot carry it.
            ParagraphProperties copy = properties == null ? new ParagraphProperties() : (ParagraphProperties)properties.CloneNode(true);
            SpacingBetweenLines spacing = copy.SpacingBetweenLines ??= new SpacingBetweenLines();
            spacing.Line = inherited!.Line!.Value;
            spacing.LineRule = inherited.LineRule?.Value ?? LineSpacingRuleValues.Auto;
            return copy;
        }
    }
}
