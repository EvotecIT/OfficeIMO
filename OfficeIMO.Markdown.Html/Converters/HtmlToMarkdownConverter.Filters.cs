using AngleSharp.Dom;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static void ApplyHtmlFilters(IElement? root, HtmlToMarkdownOptions options) {
        if (root == null || options == null) {
            return;
        }

        if (options.ExcludeSelectors.Count > 0) {
            foreach (string selector in options.ExcludeSelectors) {
                if (string.IsNullOrWhiteSpace(selector)) {
                    continue;
                }

                var matches = root.QuerySelectorAll(selector).ToList();
                for (int i = 0; i < matches.Count; i++) {
                    RemoveElement(matches[i]);
                }
            }
        }

        if (options.ElementFilters.Count == 0) {
            return;
        }

        var elements = root.QuerySelectorAll("*").ToList();
        for (int i = 0; i < elements.Count; i++) {
            var element = elements[i];
            if (element.Parent == null) {
                continue;
            }

            for (int j = 0; j < options.ElementFilters.Count; j++) {
                var filter = options.ElementFilters[j];
                if (filter != null && filter(element)) {
                    RemoveElement(element);
                    break;
                }
            }
        }
    }

    private static void RemoveElement(IElement element) {
        element.Parent?.RemoveChild(element);
    }
}
