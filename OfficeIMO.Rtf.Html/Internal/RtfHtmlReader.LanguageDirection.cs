namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private static HtmlStyleDeclaration ApplyLanguageDirectionAttributes(HtmlStyleDeclaration style, HtmlToken token) {
            RtfTextDirection? direction = null;
            if (!style.Direction.HasValue) {
                direction = HtmlStyleDeclarationParser.ParseDirection(GetAttribute(token, "dir"));
            }

            int? parsedLanguageId = null;
            string? language = GetAttribute(token, "lang");
            if (HtmlStyleDeclarationParser.TryParseLanguageId(language, out int languageId)) {
                parsedLanguageId = languageId;
            }

            if (!direction.HasValue && !parsedLanguageId.HasValue) {
                return style;
            }

            if (ReferenceEquals(style, HtmlStyleDeclaration.Empty)) {
                style = new HtmlStyleDeclaration();
            }

            if (direction.HasValue) {
                style.Direction = direction.Value;
            }

            if (parsedLanguageId.HasValue) {
                style.LanguageId = parsedLanguageId.Value;
            }

            return style;
        }

        private RtfTextDirection? ResolveTextDirection() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.Direction.HasValue) {
                    return scope.Style.Direction.Value;
                }
            }

            return null;
        }

        private int? ResolveLanguageId() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.LanguageId.HasValue) {
                    return scope.Style.LanguageId.Value;
                }
            }

            return null;
        }
    }
}
