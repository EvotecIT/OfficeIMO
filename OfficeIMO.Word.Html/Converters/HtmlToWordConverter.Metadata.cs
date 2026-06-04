using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void ApplyDocumentMetadata(WordDocument document, IDocument htmlDocument) {
            var language = GetDocumentLanguage(htmlDocument);
            if (!string.IsNullOrWhiteSpace(language)) {
                document.Settings.Language = language;
            }

            ApplyCustomDocumentProperties(document, htmlDocument);
        }

        private static string? GetDocumentLanguage(IDocument htmlDocument) {
            var root = htmlDocument.DocumentElement;
            var language = root?.GetAttribute("lang");
            if (string.IsNullOrWhiteSpace(language)) {
                language = root?.GetAttribute("xml:lang");
            }
            if (string.IsNullOrWhiteSpace(language)) {
                language = htmlDocument.Body?.GetAttribute("lang");
            }
            if (string.IsNullOrWhiteSpace(language)) {
                language = htmlDocument.Body?.GetAttribute("xml:lang");
            }

            return string.IsNullOrWhiteSpace(language) ? null : language!.Trim();
        }

        private static void ApplyCustomDocumentProperties(WordDocument document, IDocument htmlDocument) {
            foreach (var meta in htmlDocument.QuerySelectorAll("meta[data-word-custom-property]")) {
                var name = meta.GetAttribute("data-word-custom-property");
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }

                document.CustomDocumentProperties[name!.Trim()] = CreateCustomDocumentProperty(
                    meta.GetAttribute("content") ?? string.Empty,
                    meta.GetAttribute("data-property-type"));
            }
        }

        private static WordCustomProperty CreateCustomDocumentProperty(string content, string? typeName) {
            if (!Enum.TryParse<PropertyTypes>(typeName, true, out var propertyType)) {
                propertyType = PropertyTypes.Text;
            }

            return propertyType switch {
                PropertyTypes.YesNo when bool.TryParse(content, out var value) => new WordCustomProperty(value),
                PropertyTypes.DateTime when DateTime.TryParse(content, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var value) => new WordCustomProperty(value),
                PropertyTypes.NumberInteger when int.TryParse(content, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) => new WordCustomProperty(value),
                PropertyTypes.NumberDouble when double.TryParse(content, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) => new WordCustomProperty(value),
                _ => new WordCustomProperty(content)
            };
        }

        private static string? GetElementLanguage(IElement element) {
            for (IElement? current = element; current != null; current = current.ParentElement) {
                var language = current.GetAttribute("lang");
                if (string.IsNullOrWhiteSpace(language)) {
                    language = current.GetAttribute("xml:lang");
                }
                if (!string.IsNullOrWhiteSpace(language)) {
                    return language!.Trim();
                }
            }

            return null;
        }
    }
}
