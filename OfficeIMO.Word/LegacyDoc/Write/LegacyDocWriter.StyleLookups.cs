using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int MaxInheritedStyleDepth = 256;

        private static Dictionary<string, Style> IndexStyles(Styles styles) {
            var definitions = new Dictionary<string, Style>(StringComparer.OrdinalIgnoreCase);
            foreach (Style style in styles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (!string.IsNullOrWhiteSpace(id) && !definitions.ContainsKey(id!)) definitions.Add(id!, style);
            }
            return definitions;
        }
    }
}
