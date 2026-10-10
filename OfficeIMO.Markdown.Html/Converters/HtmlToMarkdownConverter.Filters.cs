using AngleSharp.Dom;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static void ApplyHtmlFilters(IElement? root, HtmlToMarkdownOptions options, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (root == null || options == null) {
            return;
        }

        if (options.ExcludeSelectors.Count > 0) {
            foreach (string selector in options.ExcludeSelectors) {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(selector)) {
                    continue;
                }

                var matches = root.QuerySelectorAll(selector).ToList();
                for (int i = 0; i < matches.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    RemoveElement(matches[i]);
                }
            }
        }

        if (options.ElementFilters.Count == 0) {
            return;
        }

        try {
            var elements = root.QuerySelectorAll("*").ToList();
            for (int i = 0; i < elements.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                var element = elements[i];
                if (element.Parent == null) {
                    continue;
                }

                for (int j = 0; j < options.ElementFilters.Count; j++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var filter = options.ElementFilters[j];
                    bool remove = filter != null && filter(OfficeIMO.Html.NativeDomBridge.Wrap(element));
                    cancellationToken.ThrowIfCancellationRequested();
                    if (remove) {
                        RemoveElement(element);
                        break;
                    }
                }
            }
        } finally {
            if (root.Owner is AngleSharp.Html.Dom.IHtmlDocument document) OfficeIMO.Html.NativeDomBridge.ReleaseCallbackSnapshot(document);
        }
    }

    private static void RemoveElement(IElement element) {
        element.Parent?.RemoveChild(element);
    }
}
