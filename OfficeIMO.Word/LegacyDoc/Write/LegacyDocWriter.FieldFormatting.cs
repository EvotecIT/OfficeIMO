using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableFormatting ReadFieldComparisonFormatting(
            Run run, LegacyDocWritableFormatting direct, LegacyDocWritableFormatting inherited) {
            // A field is flattened to one native display format. Compare the
            // inherited values before doing so, or an explicit off override can
            // erase visible content when another display run inherits hidden text.
            OpenXmlPartRootElement? root = run.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
            MainDocumentPart? main = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
            Styles? styles = main?.StyleDefinitionsPart?.Styles;
            if (main == null || styles == null) return direct.WithInheritedFormatting(inherited);

            Paragraph? paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
            string? styleId = paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                styleId = run.Ancestors<Footnote>().Any() ? "FootnoteText"
                    : run.Ancestors<Endnote>().Any() ? "EndnoteText" : "Normal";
            }
            var definitions = styles.Elements<Style>()
                .Where(style => style.Type?.Value == StyleValues.Paragraph && !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            LegacyDocWritableFormatting paragraphStyle = LegacyDocWritableFormatting.Plain;
            while (!string.IsNullOrWhiteSpace(styleId) && !string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase)
                && visited.Add(styleId!) && definitions.TryGetValue(styleId!, out Style? style)) {
                paragraphStyle = paragraphStyle.WithInheritedFormatting(ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(style)));
                styleId = style.BasedOn?.Val?.Value;
            }
            // Match the native default style's resolved fonts. Table formatting
            // supplied by the caller precedes Normal on otherwise unstyled cells.
            LegacyDocWritableFormatting normal = ReadSupportedRunFormatting(CreateDefaultParagraphStyle(main, definitions, styles).StyleRunProperties);
            return direct.WithInheritedFormatting(paragraphStyle.WithInheritedFormatting(inherited).WithInheritedFormatting(normal));
        }
    }
}
