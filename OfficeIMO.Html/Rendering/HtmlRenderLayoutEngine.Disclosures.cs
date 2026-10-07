using AngleSharp.Dom;
using System.Text;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HtmlDisclosureState _disclosures = new HtmlDisclosureState();

    /// <summary>
    /// Closed disclosures expose their first summary element only. This filters
    /// direct text as well as elements, before layout or intrinsic measurement;
    /// authored child display values cannot open the disclosure's content slot.
    /// </summary>
    private bool IsClosedDisclosureChild(INode node) => _disclosures.IsClosedChild(node);

    /// <summary>Applies the same ancestor disclosure state to independently collected bookmark text.</summary>
    private bool IsInsideClosedDisclosure(INode node) => _disclosures.IsInsideClosedContent(node);

    /// <summary>
    /// Keeps the existing text-content measurement for table and shrink-to-fit
    /// widths while excluding closed content. Traversal uses layout limits and
    /// cancellation, and preserves visible artifact text used for sizing.
    /// </summary>
    private string ResolveDisclosureTextContent(IElement element, int depth) {
        var text = new StringBuilder();
        var pending = new Stack<(INode Node, int Depth)>();
        pending.Push((element, depth));
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        while (pending.Count > 0) {
            CheckCancellation();
            (INode node, int nodeDepth) = pending.Pop();
            if (IsClosedDisclosureChild(node)) continue;
            ChargeLayoutOperation(source);
            if (node is IText value) {
                text.Append(value.Data);
            } else if (node is IElement childElement) {
                EnsureDepth(nodeDepth, childElement);
                for (INode? child = childElement.LastChild; child != null; child = child.PreviousSibling) {
                    pending.Push((child, nodeDepth + 1));
                }
            }
        }
        return text.ToString();
    }
}
