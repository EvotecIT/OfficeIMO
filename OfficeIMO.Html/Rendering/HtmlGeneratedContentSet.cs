using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlGeneratedContentSet {
    private readonly IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> _content;

    internal HtmlGeneratedContentSet(IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> content) {
        _content = content ?? throw new ArgumentNullException(nameof(content));
    }

    internal bool TryGet(IElement element, HtmlPseudoElementKind kind, out string content) {
        if (TryGetContent(element, kind, out HtmlGeneratedContent found)
            && !string.IsNullOrEmpty(found.Text)) {
            content = found.Text;
            return true;
        }

        content = string.Empty;
        return false;
    }

    internal bool TryGetContent(IElement element, HtmlPseudoElementKind kind, out HtmlGeneratedContent content) {
        if (_content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)) {
            HtmlGeneratedContent? found = kind switch {
                HtmlPseudoElementKind.Before => pair.Before,
                HtmlPseudoElementKind.After => pair.After,
                _ => pair.Marker
            };
            if (found != null && found.Fragments.Count > 0) {
                content = found;
                return true;
            }
        }

        content = null!;
        return false;
    }

    internal bool Suppresses(IElement element, HtmlPseudoElementKind kind) =>
        kind == HtmlPseudoElementKind.Marker
        && _content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)
        && pair.SuppressMarker;
}

internal sealed class HtmlGeneratedPseudoContentPair {
    internal HtmlGeneratedContent? Before { get; set; }
    internal HtmlGeneratedContent? After { get; set; }
    internal HtmlGeneratedContent? Marker { get; set; }
    internal bool SuppressMarker { get; set; }
}

internal sealed class HtmlGeneratedContent {
    internal HtmlGeneratedContent(IReadOnlyList<HtmlGeneratedContentFragment> fragments) {
        Fragments = fragments ?? throw new ArgumentNullException(nameof(fragments));
        Text = string.Concat(fragments.Where(fragment => fragment.Kind == HtmlGeneratedContentFragmentKind.Text).Select(fragment => fragment.Value));
    }

    internal IReadOnlyList<HtmlGeneratedContentFragment> Fragments { get; }
    internal string Text { get; }
}

internal enum HtmlGeneratedContentFragmentKind {
    Text,
    Image
}

internal readonly record struct HtmlGeneratedContentFragment(HtmlGeneratedContentFragmentKind Kind, string Value);
