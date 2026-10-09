using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Finds the nearest transformed padding box within the active layout root.
    /// A fixed descendant of that box uses local positioning and the box's paint
    /// effects; only fixed descendants without such a box use the viewport.
    /// </summary>
    private IElement? ResolveTransformedContainingBlock(IElement? directParent, HtmlRenderBoxStyle? directParentStyle = null) {
        double referenceWidth = Math.Max(1D, ActiveSurfaceWidth - ActiveMargins.Left - ActiveMargins.Right);
        for (IElement? ancestor = directParent; ancestor != null; ancestor = ancestor.ParentElement) {
            HtmlRenderBoxStyle style = ReferenceEquals(ancestor, directParent) && directParentStyle != null
                ? directParentStyle
                : _styleResolver.Resolve(ancestor, referenceWidth);
            if (EstablishesTransformedContainingBlock(style)) return ancestor;
            if (IsRootLayoutContainer(ancestor)) break;
        }
        return null;
    }

    // Non-replaced inline boxes and display:contents have no transformed box
    // to establish a containing block (CSS Transforms 1, sections 2 and 3).
    private static bool EstablishesTransformedContainingBlock(HtmlRenderBoxStyle style) =>
        style.Transform != "none"
        && style.Display is not ("inline" or "contents" or "none" or "table-column" or "table-column-group");
}
