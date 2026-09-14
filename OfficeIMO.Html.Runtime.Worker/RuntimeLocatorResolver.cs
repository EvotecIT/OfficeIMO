using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeLocatorResolver(IDocument document, HtmlScriptRequest options) {
    internal IReadOnlyList<IElement> Resolve(HtmlLocatorQuery query, CancellationToken token) {
        // Bound the tree before provider selector evaluation, including nonmatching nodes.
        var ordered = Enumerate(token).Select(item => item.Element).ToArray();
        return ResolveQuery(query, ordered, new HtmlAccessibilitySemantics.HtmlAccessibleNameContext(), token);
    }

    internal IReadOnlyList<RuntimeDocumentElement> Enumerate(CancellationToken token) {
        var ordered = new List<RuntimeDocumentElement>();
        var pending = new Stack<(INode Node, int Depth)>();
        pending.Push((document, 0));
        int nodes = 0;
        while (pending.Count > 0) {
            token.ThrowIfCancellationRequested();
            var current = pending.Pop();
            if (++nodes > options.MaxNodes || current.Depth > options.MaxDepth)
                throw new HtmlScriptRuntimeException("The automation document exceeds its node or depth budget.");
            if (current.Node is IElement element) {
                int? parentIndex = null;
                for (INode? parent = element.Parent; parent != null; parent = parent.Parent) {
                    if (parent is not IElement parentElement) continue;
                    for (int index = ordered.Count - 1; index >= 0; index--)
                        if (ReferenceEquals(ordered[index].Element, parentElement)) { parentIndex = index; break; }
                    break;
                }
                ordered.Add(new RuntimeDocumentElement(element, current.Depth, parentIndex));
            }
            for (int i = current.Node.ChildNodes.Length - 1; i >= 0; i--)
                pending.Push((current.Node.ChildNodes[i], current.Depth + (current.Node.ChildNodes[i] is IElement ? 1 : 0)));
        }
        return ordered;
    }

    internal IElement? Resolve(HtmlObservedElementReference reference, string pageId, long revision, CancellationToken token) {
        if (!string.Equals(reference.PageId, pageId, StringComparison.Ordinal) || reference.Revision != revision) return null;
        IReadOnlyList<RuntimeDocumentElement> ordered = Enumerate(token);
        if (reference.ElementIndex >= ordered.Count) return null;
        IElement element = ordered[reference.ElementIndex].Element;
        return string.Equals(element.LocalName, reference.ElementName, StringComparison.Ordinal)
            && string.Equals(element.Id ?? string.Empty, reference.ElementId, StringComparison.Ordinal) ? element : null;
    }

    private IReadOnlyList<IElement> ResolveQuery(HtmlLocatorQuery query, IReadOnlyList<IElement> ordered, HtmlAccessibilitySemantics.HtmlAccessibleNameContext names, CancellationToken token) {
        IEnumerable<IParentNode> roots = query.Scope == null ? new[] { (IParentNode)document }
            : ResolveQuery(query.Scope, ordered, names, token).Cast<IParentNode>();
        var matches = new HashSet<IElement>();
        string value = Normalize(query.Value);
        foreach (var root in roots) {
            token.ThrowIfCancellationRequested();
            foreach (IElement element in root.QuerySelectorAll(query.Kind == HtmlLocatorKind.Css ? query.Value : "*")) {
                token.ThrowIfCancellationRequested();
                if (query.Kind == HtmlLocatorKind.Css) { matches.Add(element); continue; }
                if (query.Kind == HtmlLocatorKind.Text && element.LocalName is "script" or "style" or "template") continue;
                string candidate = query.Kind == HtmlLocatorKind.Text ? Normalize(element.TextContent)
                    : Normalize(AccessibleName(element, names));
                bool Match(string text) => query.Exact ? string.Equals(text, value, StringComparison.Ordinal) : text.Contains(value, StringComparison.Ordinal);
                if (!Match(candidate)) continue;
                if (query.Kind == HtmlLocatorKind.Text && element.Children.Any(child => Match(Normalize(child.TextContent)))) continue;
                matches.Add(element);
            }
        }
        IElement[] result = ordered.Where(matches.Contains).ToArray();
        return query.Index is int index ? index < result.Length ? new[] { result[index] } : Array.Empty<IElement>() : result;
    }

    internal static string Normalize(string? value) => string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    internal static string AccessibleName(IElement element, HtmlAccessibilitySemantics.HtmlAccessibleNameContext? context = null) => HtmlAccessibilitySemantics.GetAccessibleName(element,
        includeTextFallback: element.LocalName is "button" or "a" or "summary" or "option" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6"
            || element.HasAttribute("role"), context ?? new HtmlAccessibilitySemantics.HtmlAccessibleNameContext());
}


internal readonly record struct RuntimeDocumentElement(IElement Element, int Depth, int? ParentElementIndex);
