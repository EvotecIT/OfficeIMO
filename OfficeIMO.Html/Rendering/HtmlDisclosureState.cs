using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>Shares the first-summary content slot rule across one render operation.</summary>
internal sealed class HtmlDisclosureState {
    private readonly Dictionary<IElement, IElement?> _summaries = new Dictionary<IElement, IElement?>();

    internal bool IsClosedChild(INode node) {
        IElement? parent = node.ParentElement;
        if (parent == null || !string.Equals(parent.LocalName, "details", StringComparison.OrdinalIgnoreCase)
            || parent.HasAttribute("open")) return false;
        if (!_summaries.TryGetValue(parent, out IElement? summary)) {
            summary = parent.Children.FirstOrDefault(child => string.Equals(child.LocalName, "summary", StringComparison.OrdinalIgnoreCase));
            _summaries.Add(parent, summary);
        }
        return !ReferenceEquals(node, summary);
    }

    internal bool IsInsideClosedContent(INode node) {
        for (INode? current = node; current?.ParentElement != null; current = current.ParentElement) {
            if (IsClosedChild(current)) return true;
        }
        return false;
    }
}
