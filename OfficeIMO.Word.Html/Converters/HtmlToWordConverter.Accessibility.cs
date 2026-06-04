using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly HashSet<string> _weakLinkTexts = new(StringComparer.OrdinalIgnoreCase) {
            "click here",
            "continue",
            "details",
            "go",
            "here",
            "learn more",
            "link",
            "more",
            "read more"
        };

        private int _lastAccessibilityHeadingLevel;

        private void ResetAccessibilityDiagnosticsState() {
            _lastAccessibilityHeadingLevel = 0;
        }

        private void ReportAccessibilityDiagnostics(IElement element) {
            if (!_options.EnableAccessibilityDiagnostics) {
                return;
            }

            var tagName = element.TagName.ToLowerInvariant();
            switch (tagName) {
                case "img":
                    ReportImageAccessibility((IHtmlImageElement)element);
                    break;
                case "a":
                    ReportLinkAccessibility(element);
                    break;
                case "table":
                    ReportTableAccessibility(element);
                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    ReportHeadingAccessibility(tagName);
                    break;
            }
        }

        private void ReportImageAccessibility(IHtmlImageElement image) {
            if (IsHiddenFromAccessibility(image) || image.HasAttribute("alt")) {
                return;
            }

            AddDiagnostic(
                _options,
                "AccessibilityImageMissingAlt",
                "Image is missing an alt attribute. Add meaningful alternate text, or use an empty alt value for decorative images.",
                GetElementSource(image));
        }

        private void ReportLinkAccessibility(IElement link) {
            var href = link.GetAttribute("href");
            if (string.IsNullOrWhiteSpace(href) || IsHiddenFromAccessibility(link)) {
                return;
            }
            var hrefSource = href!.Trim();

            var accessibleName = GetAccessibleName(link);
            if (string.IsNullOrWhiteSpace(accessibleName)) {
                AddDiagnostic(
                    _options,
                    "AccessibilityLinkTextMissing",
                    "Link has no accessible text. Add visible link text, aria-label, title, or image alternate text.",
                    hrefSource);
                return;
            }

            if (IsWeakLinkText(accessibleName, hrefSource)) {
                AddDiagnostic(
                    _options,
                    "AccessibilityLinkTextWeak",
                    "Link text is generic or URL-only. Use text that describes the destination or action.",
                    hrefSource);
            }
        }

        private void ReportHeadingAccessibility(string tagName) {
            var level = tagName.Length == 2 && char.IsDigit(tagName[1]) ? tagName[1] - '0' : 0;
            if (level <= 0) {
                return;
            }

            if (_lastAccessibilityHeadingLevel > 0 && level > _lastAccessibilityHeadingLevel + 1) {
                AddDiagnostic(
                    _options,
                    "AccessibilityHeadingLevelSkipped",
                    $"Heading level jumps from h{_lastAccessibilityHeadingLevel} to h{level}. Use consecutive heading levels where practical.",
                    tagName);
            }

            _lastAccessibilityHeadingLevel = level;
        }

        private void ReportTableAccessibility(IElement table) {
            if (!IsLikelyDataTable(table) || HasTableHeader(table)) {
                return;
            }

            AddDiagnostic(
                _options,
                "AccessibilityTableMissingHeader",
                "Data table has no header cells. Use th, thead, or scope attributes so exported documents retain table meaning.",
                GetElementSource(table));
        }

        private static bool IsHiddenFromAccessibility(IElement element) {
            var ariaHidden = element.GetAttribute("aria-hidden");
            if (string.Equals(ariaHidden, "true", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            var role = element.GetAttribute("role");
            return string.Equals(role, "presentation", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetAccessibleName(IElement element) {
            var ariaLabel = element.GetAttribute("aria-label");
            if (!string.IsNullOrWhiteSpace(ariaLabel)) {
                return NormalizeAccessibleText(ariaLabel);
            }

            var text = NormalizeAccessibleText(element.TextContent);
            if (!string.IsNullOrWhiteSpace(text)) {
                return text;
            }

            foreach (var image in element.QuerySelectorAll("img")) {
                var alt = image.GetAttribute("alt");
                if (!string.IsNullOrWhiteSpace(alt)) {
                    return NormalizeAccessibleText(alt);
                }
            }

            var title = element.GetAttribute("title");
            return NormalizeAccessibleText(title);
        }

        private static bool IsWeakLinkText(string accessibleName, string href) {
            var normalized = NormalizeAccessibleText(accessibleName).Trim('.', ':');
            if (_weakLinkTexts.Contains(normalized)) {
                return true;
            }

            if (Uri.TryCreate(normalized, UriKind.Absolute, out _)) {
                return true;
            }

            return string.Equals(normalized, href, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeAccessibleText(string? text) {
            if (string.IsNullOrWhiteSpace(text)) {
                return string.Empty;
            }

            return string.Join(" ", text!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool IsLikelyDataTable(IElement table) {
            if (string.Equals(table.GetAttribute("role"), "presentation", StringComparison.OrdinalIgnoreCase)
                || string.Equals(table.GetAttribute("role"), "none", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(table.GetAttribute("summary"))
                || table.QuerySelector("caption") != null
                || table.QuerySelector("thead") != null
                || table.QuerySelector("th") != null
                || table.QuerySelector("[scope]") != null) {
                return true;
            }

            return table.QuerySelectorAll("tr").Length > 1
                && table.QuerySelectorAll("td").Length > 1;
        }

        private static bool HasTableHeader(IElement table) =>
            table.QuerySelector("th") != null
            || table.QuerySelector("thead") != null
            || table.QuerySelector("[scope]") != null;

        private static string GetElementSource(IElement element) {
            var id = element.GetAttribute("id");
            if (!string.IsNullOrWhiteSpace(id)) {
                return element.TagName.ToLowerInvariant() + "#" + id;
            }

            var src = element.GetAttribute("src");
            if (!string.IsNullOrWhiteSpace(src)) {
                return src!;
            }

            return element.TagName.ToLowerInvariant();
        }
    }
}
