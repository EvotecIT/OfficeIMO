using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlGeneratedContentSet {
    private readonly IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> _content;

    internal HtmlGeneratedContentSet(IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> content) {
        _content = content ?? throw new ArgumentNullException(nameof(content));
    }

    internal bool TryGet(IElement element, HtmlPseudoElementKind kind, out string content) {
        if (_content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)) {
            string? found = kind switch {
                HtmlPseudoElementKind.Before => pair.Before,
                HtmlPseudoElementKind.After => pair.After,
                _ => pair.Marker
            };
            if (!string.IsNullOrEmpty(found)) {
                content = found!;
                return true;
            }
        }

        content = string.Empty;
        return false;
    }

    internal bool Suppresses(IElement element, HtmlPseudoElementKind kind) =>
        kind == HtmlPseudoElementKind.Marker
        && _content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)
        && pair.SuppressMarker;
}

internal sealed class HtmlGeneratedPseudoContentPair {
    internal string? Before { get; set; }
    internal string? After { get; set; }
    internal string? Marker { get; set; }
    internal bool SuppressMarker { get; set; }
}
