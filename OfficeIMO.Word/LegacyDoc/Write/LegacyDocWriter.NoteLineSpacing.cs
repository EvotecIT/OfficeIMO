using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static ParagraphProperties? MaterializeNoteLineSpacing(Paragraph paragraph, string noteKind) {
            ParagraphProperties? properties = paragraph.ParagraphProperties;
            if (!string.IsNullOrWhiteSpace(properties?.SpacingBetweenLines?.Line?.Value)) return properties;
            OpenXmlPartRootElement? root = paragraph.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
            Styles? styles = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return properties;
            string? styleId = properties?.ParagraphStyleId?.Val?.Value ?? (noteKind == "footnote" ? "FootnoteText" : "EndnoteText");
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            SpacingBetweenLines? inherited = null;
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                Style? style = styles.Elements<Style>().FirstOrDefault(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
                inherited = style?.StyleParagraphProperties?.SpacingBetweenLines;
                if (!string.IsNullOrWhiteSpace(inherited?.Line?.Value)) break;
                styleId = style?.BasedOn?.Val?.Value;
            }
            if (string.IsNullOrWhiteSpace(inherited?.Line?.Value)) inherited = styles.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
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
