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

        private void ApplyDocumentLanguageDirection(string name, HtmlStyleDeclaration style) {
            if (!string.Equals(name, "html", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(name, "body", StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            if (style.LanguageId.HasValue) {
                _document.Settings.DefaultLanguageId = style.LanguageId.Value;
            }

            if (style.Direction.HasValue) {
                _document.Settings.Direction = style.Direction.Value;
            }
        }

        private RtfTextDirection? ResolveTextDirection() {
            foreach (HtmlStyleScope scope in _styles) {
                if (IsBlockDirectionScope(scope.Name)) {
                    continue;
                }

                if (scope.Style.Direction.HasValue) {
                    return scope.Style.Direction.Value;
                }
            }

            return null;
        }

        private int? ResolveLanguageId() {
            foreach (HtmlStyleScope scope in _styles) {
                if (IsDocumentScope(scope.Name)) {
                    continue;
                }

                if (scope.Style.LanguageId.HasValue) {
                    return scope.Style.LanguageId.Value;
                }
            }

            return null;
        }

        private static bool IsDocumentScope(string name) =>
            string.Equals(name, "html", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "body", StringComparison.OrdinalIgnoreCase);

        private static bool IsBlockDirectionScope(string name) {
            switch (name) {
                case "html":
                case "body":
                case "p":
                case "div":
                case "section":
                case "article":
                case "blockquote":
                case "li":
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    return true;
                default:
                    return false;
            }
        }
    }
}
