using AngleSharp.Css.Dom;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void ApplyPageBreakAfterFromCss(WordParagraph paragraph, IElement element) {
            var styleAttribute = element.GetAttribute("style");
            var declaration = ParseInlineDeclaration(styleAttribute);
            if (StyleRequestsPageBreakAfter(declaration, styleAttribute) || StyleTextRequestsPageBreak(styleAttribute, "break-after", "page-break-after")) {
                AddPageBreakAfter(paragraph);
            }
        }

        private static void AddPageBreakAfter(WordParagraph paragraph) {
            paragraph.AddParagraphAfterSelf().AddBreak(BreakValues.Page);
        }

        private static bool StyleRequestsPageBreakBefore(IElement element) {
            var styleAttribute = element.GetAttribute("style");
            var declaration = ParseInlineDeclaration(styleAttribute);
            return StyleRequestsPageBreakBefore(declaration, styleAttribute) || StyleTextRequestsPageBreak(styleAttribute, "break-before", "page-break-before");
        }

        private static bool StyleRequestsPageBreakAfter(IElement element) {
            var styleAttribute = element.GetAttribute("style");
            var declaration = ParseInlineDeclaration(styleAttribute);
            return StyleRequestsPageBreakAfter(declaration, styleAttribute) || StyleTextRequestsPageBreak(styleAttribute, "break-after", "page-break-after");
        }

        private static bool StyleRequestsPageBreakBefore(ICssStyleDeclaration declaration, string? styleAttribute) =>
            IsPageBreakValue(GetInlinePropertyValue(declaration, styleAttribute, "break-before")) ||
            IsPageBreakValue(GetInlinePropertyValue(declaration, styleAttribute, "page-break-before"));

        private static bool StyleRequestsPageBreakAfter(ICssStyleDeclaration declaration, string? styleAttribute) =>
            IsPageBreakValue(GetInlinePropertyValue(declaration, styleAttribute, "break-after")) ||
            IsPageBreakValue(GetInlinePropertyValue(declaration, styleAttribute, "page-break-after"));

        private static bool StyleTextRequestsPageBreak(string? styleAttribute, params string[] propertyNames) {
            if (string.IsNullOrWhiteSpace(styleAttribute)) {
                return false;
            }

            var styleText = styleAttribute!;
            foreach (var declaration in styleText.Split(';')) {
                var separatorIndex = declaration.IndexOf(':');
                if (separatorIndex < 0) {
                    continue;
                }

                var name = declaration.Substring(0, separatorIndex).Trim();
                if (!propertyNames.Any(propertyName => string.Equals(propertyName, name, StringComparison.OrdinalIgnoreCase))) {
                    continue;
                }

                var value = declaration.Substring(separatorIndex + 1).Trim();
                var importantIndex = value.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
                if (importantIndex >= 0) {
                    value = value.Substring(0, importantIndex).Trim();
                }

                if (IsPageBreakValue(value)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsPageBreakProperty(string name) =>
            string.Equals(name, "break-before", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "break-after", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "page-break-before", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "page-break-after", StringComparison.OrdinalIgnoreCase);

        private static bool IsPageBreakValue(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            return value!.Trim().ToLowerInvariant() switch {
                "always" => true,
                "page" => true,
                "left" => true,
                "right" => true,
                "recto" => true,
                "verso" => true,
                _ => false,
            };
        }
    }
}
