using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlGeneratedContentSet {
    private readonly IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> _content;
    private readonly IReadOnlyDictionary<string, int> _targetPages;

    internal HtmlGeneratedContentSet(
        IReadOnlyDictionary<IElement, HtmlGeneratedPseudoContentPair> content,
        IReadOnlyDictionary<string, int>? targetPages = null) {
        _content = content ?? throw new ArgumentNullException(nameof(content));
        _targetPages = targetPages ?? new Dictionary<string, int>(StringComparer.Ordinal);
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
                HtmlPseudoElementKind.Marker => pair.Marker,
                HtmlPseudoElementKind.FootnoteCall => pair.FootnoteCall,
                _ => pair.FootnoteMarker
            };
            if (found != null && found.Fragments.Count > 0) {
                content = found;
                return true;
            }
        }

        content = null!;
        return false;
    }

    internal bool Suppresses(IElement element, HtmlPseudoElementKind kind) {
        if (!_content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)) return false;
        if (kind == HtmlPseudoElementKind.Marker) return pair.SuppressMarker;
        if (kind == HtmlPseudoElementKind.FootnoteCall) return pair.SuppressFootnoteCall;
        return kind == HtmlPseudoElementKind.FootnoteMarker && pair.SuppressFootnoteMarker;
    }

    internal bool HasTargetPageReferences => _content.Values.Any(pair =>
        ContainsTargetPage(pair.Before) || ContainsTargetPage(pair.After) || ContainsTargetPage(pair.Marker)
        || ContainsTargetPage(pair.FootnoteCall) || ContainsTargetPage(pair.FootnoteMarker));

    internal IReadOnlyCollection<string> TargetPageIds => _content.Values
        .SelectMany(pair => EnumerateTargetPageIds(pair.Before)
            .Concat(EnumerateTargetPageIds(pair.After))
            .Concat(EnumerateTargetPageIds(pair.Marker))
            .Concat(EnumerateTargetPageIds(pair.FootnoteCall))
            .Concat(EnumerateTargetPageIds(pair.FootnoteMarker)))
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    internal HtmlGeneratedContentSet WithTargetPages(IReadOnlyDictionary<string, int> targetPages) =>
        new HtmlGeneratedContentSet(_content, targetPages);

    internal bool TryGetTargetPage(string id, out int pageNumber) => _targetPages.TryGetValue(id, out pageNumber);

    internal bool TargetPagesEqual(IReadOnlyDictionary<string, int> targetPages) =>
        _targetPages.Count == targetPages.Count
        && _targetPages.All(pair => targetPages.TryGetValue(pair.Key, out int page) && page == pair.Value);

    private static bool ContainsTargetPage(HtmlGeneratedContent? content) =>
        content != null && content.Fragments.Any(fragment => fragment.Kind == HtmlGeneratedContentFragmentKind.TargetPage);

    private static IEnumerable<string> EnumerateTargetPageIds(HtmlGeneratedContent? content) =>
        content?.Fragments
            .Where(fragment => fragment.Kind == HtmlGeneratedContentFragmentKind.TargetPage)
            .Select(fragment => fragment.Value)
        ?? Enumerable.Empty<string>();
}

internal sealed class HtmlGeneratedPseudoContentPair {
    internal HtmlGeneratedContent? Before { get; set; }
    internal HtmlGeneratedContent? After { get; set; }
    internal HtmlGeneratedContent? Marker { get; set; }
    internal HtmlGeneratedContent? FootnoteCall { get; set; }
    internal HtmlGeneratedContent? FootnoteMarker { get; set; }
    internal bool SuppressMarker { get; set; }
    internal bool SuppressFootnoteCall { get; set; }
    internal bool SuppressFootnoteMarker { get; set; }
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
    Image,
    Leader,
    TargetPage
}

internal readonly record struct HtmlGeneratedContentFragment(
    HtmlGeneratedContentFragmentKind Kind,
    string Value,
    string? Format = null);
