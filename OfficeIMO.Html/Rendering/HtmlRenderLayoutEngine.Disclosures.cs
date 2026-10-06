using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, IElement?> _disclosureSummaries = new Dictionary<IElement, IElement?>();

    /// <summary>
    /// Closed disclosures expose their first summary element only. This filters
    /// direct text as well as elements, before layout or intrinsic measurement;
    /// authored child display values cannot open the disclosure's content slot.
    /// </summary>
    private bool IsClosedDisclosureChild(INode node) {
        IElement? parent = node.ParentElement;
        if (parent == null || !string.Equals(parent.LocalName, "details", StringComparison.OrdinalIgnoreCase)
            || parent.HasAttribute("open")) return false;
        if (!_disclosureSummaries.TryGetValue(parent, out IElement? summary)) {
            summary = parent.Children.FirstOrDefault(child => string.Equals(child.LocalName, "summary", StringComparison.OrdinalIgnoreCase));
            _disclosureSummaries.Add(parent, summary);
        }
        return !ReferenceEquals(node, summary);
    }

    /// <summary>Applies the same ancestor disclosure state to independently collected bookmark text.</summary>
    private bool IsInsideClosedDisclosure(INode node) {
        for (INode? current = node; current?.ParentElement != null; current = current.ParentElement) {
            if (IsClosedDisclosureChild(current)) return true;
        }
        return false;
    }
}
