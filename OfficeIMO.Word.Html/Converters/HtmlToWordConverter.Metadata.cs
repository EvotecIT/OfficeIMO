using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void ApplyDocumentMetadata(WordDocument document, IDocument htmlDocument) {
            var language = GetDocumentLanguage(htmlDocument);
            if (!string.IsNullOrWhiteSpace(language)) {
                document.Settings.Language = language;
            }
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
