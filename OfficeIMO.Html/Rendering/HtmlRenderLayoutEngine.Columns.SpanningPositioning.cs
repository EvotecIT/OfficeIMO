using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, IElement> _columnSpannerContainingOwners = new();

    /// <summary>
    /// A spanner's layout containing-block ancestry resumes at its multicol
    /// owner. Its original DOM ancestry still owns styles and semantics.
    /// Retain this relation until deferred positioned descendants are laid out.
    /// </summary>
    private void RegisterColumnSpannerContainingOwner(IElement spanner, IElement owner) {
        CheckCancellation();
        ChargeLayoutOperation("column spanner containing-block ancestry");
        _columnSpannerContainingOwners[spanner] = owner;
    }

    /// <summary>Advances the existing containing-block resolvers through the effective layout ancestry.</summary>
    private IElement? ResolveColumnContainingParent(IElement element) =>
        _columnSpannerContainingOwners.TryGetValue(element, out IElement? owner)
            ? owner : element.ParentElement;

    private static bool IsColumnSourceTopAnchored(HtmlRenderBoxStyle style) =>
        style.Position == "absolute" && !IsAutoInset(style.Top)
        && (!IsAutoInset(style.Left) || !IsAutoInset(style.Right));

    /// <summary>Retains a deferred positioned request's original source-piece ownership during repeated layout.</summary>
    private bool IsColumnPositionedRequestSelected(PositionedElementRequest request) {
        if (!TryGetColumnSourceFragment(request.ContainingBlock, out ColumnSourceFragment containing)) return true;
        if (_columnSourceFlowContext?.FirstFragmentPositionedElements.Contains(request.Element) == true)
            return containing.FirstFragment;
        for (IElement current = request.Element; current.ParentElement is { } parent; current = parent) {
            CheckCancellation();
            ChargeLayoutOperation("column positioned source selection");
            if (TryGetColumnSourceFragment(parent, out ColumnSourceFragment fragment)) {
                bool selected = false;
                foreach (INode node in fragment.Nodes) {
                    CheckCancellation();
                    ChargeLayoutOperation("column positioned selected child");
                    if (!ReferenceEquals(node, current)) continue;
                    selected = true;
                    break;
                }
                if (!selected) return false;
            }
            if (ReferenceEquals(parent, request.ContainingBlock)) return true;
        }
        return true;
    }

    /// <summary>
    /// Discovers explicitly anchored children before their containing ancestor's
    /// first fragment is painted. Automatic insets still need their actual
    /// source-flow anchor and retain normal discovery.
    /// </summary>
    private void PrepareColumnSourcePositionedRequests(IReadOnlyList<MultiColumnSpanPartition> partitions) {
        var visited = new HashSet<IElement>();
        foreach (MultiColumnSpanPartition partition in partitions) {
            foreach (ColumnSpannerAncestor ancestor in partition.SpannerAncestors) {
                PrepareChildren(ancestor.Element, ancestor.Style, ancestor.ContainingWidth, ancestor.Depth);
            }
        }

        void PrepareChildren(IElement container, HtmlRenderBoxStyle style, double containingWidth, int depth) {
            CheckCancellation();
            EnsureDepth(depth, container);
            if (!visited.Add(container)) return;
            ChargeLayoutOperation("column source positioned discovery");
            double width = style.Display is "inline" or "contents" ? containingWidth
                : Math.Max(1D, ResolveBoxWidth(Math.Max(1D,
                    containingWidth - style.MarginLeft - style.MarginRight), style) - style.HorizontalInsets);
            foreach (IElement child in container.Children) {
                CheckCancellation();
                ChargeLayoutOperation("column source positioned child");
                if (ShouldSkipElement(child) || IsClosedDisclosureChild(child)) continue;
                HtmlRenderBoxStyle childStyle = ForwardContainingHeightBasis(child,
                    _styleResolver.Resolve(child, width, style), style);
                if (childStyle.Display == "none") continue;
                if (ShouldExtractOutOfFlow(childStyle)) {
                    if (IsColumnSourceTopAnchored(childStyle)) {
                        RegisterOutOfFlowElement(container, child, childStyle, style, depth + 1);
                        _columnSourceFlowContext?.FirstFragmentPositionedElements.Add(child);
                    }
                    continue;
                }
                if (IsColumnSourceSpanner(child, childStyle) || !CanTraverseColumnSpannerSource(child, childStyle)) continue;
                childStyle = ResolveOrdinaryIntrinsicWidths(child, childStyle, width, depth + 1);
                PrepareChildren(child, childStyle, width, depth + 1);
            }
        }
    }
}
