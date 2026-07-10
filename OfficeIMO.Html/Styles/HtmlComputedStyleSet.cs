using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal enum HtmlPseudoElementKind {
    Before,
    After
}

internal sealed class HtmlComputedStyleSet {
    private readonly IReadOnlyDictionary<IElement, HtmlPseudoElementStylePair> _pseudoElements;

    internal HtmlComputedStyleSet(
        IReadOnlyDictionary<IElement, HtmlComputedStyle> elements,
        IReadOnlyDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements) {
        Elements = elements ?? throw new ArgumentNullException(nameof(elements));
        _pseudoElements = pseudoElements ?? throw new ArgumentNullException(nameof(pseudoElements));
    }

    internal IReadOnlyDictionary<IElement, HtmlComputedStyle> Elements { get; }

    internal bool HasPseudoElements => _pseudoElements.Count > 0;

    internal bool TryGetPseudoStyle(IElement element, HtmlPseudoElementKind kind, out HtmlComputedStyle style) {
        if (_pseudoElements.TryGetValue(element, out HtmlPseudoElementStylePair? pair)) {
            HtmlComputedStyle? found = kind == HtmlPseudoElementKind.Before ? pair.Before : pair.After;
            if (found != null) {
                style = found;
                return true;
            }
        }

        style = null!;
        return false;
    }
}

internal sealed class HtmlPseudoElementStylePair {
    internal HtmlComputedStyle? Before { get; set; }
    internal HtmlComputedStyle? After { get; set; }
}
