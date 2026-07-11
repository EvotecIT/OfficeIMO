using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlGeneratedContentSet {
    private readonly IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> _content;

    internal HtmlGeneratedContentSet(IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> content) {
        _content = content ?? throw new ArgumentNullException(nameof(content));
    }

    internal bool TryGet(IElement element, HtmlPseudoElementKind kind, out string content) {
        if (_content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)) {
            string? found = kind == HtmlPseudoElementKind.Before ? pair.Before : pair.After;
            if (!string.IsNullOrEmpty(found)) {
                content = found!;
                return true;
            }
        }

        content = string.Empty;
        return false;
    }
}

internal sealed class HtmlGeneratedPseudoContentPair {
    internal string? Before { get; set; }
    internal string? After { get; set; }
}
