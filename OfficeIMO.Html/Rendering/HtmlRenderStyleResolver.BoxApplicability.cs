namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    // Classify the original used box before specialized layout owners reset
    // position/float or clone a blockified flex/grid item's style.
    private bool IsTransformableBox(IElement element, HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle? parent, bool pseudoElement) {
        if (style.Display is "none" or "contents") return false;
        if (IsBlockifiedBox(element, style, parent, pseudoElement)) return true;
        if (style.Display is "table-column" or "table-column-group") return false;
        return style.Display != "inline" || IsReplacedOrAtomicInline(element, pseudoElement);
    }

    private bool AcceptsBoxDimensions(IElement element, HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle? parent, bool pseudoElement) =>
        style.Display is not ("none" or "contents")
        && (style.Display != "inline" || IsBlockifiedBox(element, style, parent, pseudoElement)
            || IsReplacedOrAtomicInline(element, pseudoElement));

    private bool IsBlockifiedBox(IElement element, HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle? parent, bool pseudoElement) {
        if (style.Position is "absolute" or "fixed" || style.FloatSide != "none") return true;
        if (!pseudoElement && ReferenceEquals(element.Owner?.DocumentElement, element)) return true;
        // A pseudo's supplied parent is its originating element. Parentless
        // ancestor lookups recover the same context directly from computed CSS.
        IElement? ancestor = pseudoElement ? element : element.ParentElement;
        string parentDisplay = parent?.Display ?? ResolveElementDisplay(ancestor);
        while (parentDisplay == "contents" && ancestor?.ParentElement != null) {
            ancestor = ancestor.ParentElement;
            parentDisplay = ResolveElementDisplay(ancestor);
        }
        return parentDisplay is "flex" or "inline-flex" or "grid" or "inline-grid";
    }

    private string ResolveElementDisplay(IElement? element) => element == null ? string.Empty
        : _computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)
            ? ResolveDisplay(element, computed)
            : HtmlElementDisplay.GetDefaultValue(element);

    // Normal layout and parentless recovery need the same legacy clamp context.
    private static string ResolveDisplay(IElement element, HtmlComputedStyle computed) =>
        ResolveDisplay(element, computed.GetValue("display"), computed.GetValue("-webkit-box-orient"),
            ResolveLineClamp(computed).HasValue);

    // These existing specialized owners lay out atomic boxes even where their
    // HTML display default is inline. Generated text on a replaced origin is
    // still a separate non-replaced pseudo box.
    private static bool IsReplacedOrAtomicInline(IElement element, bool pseudoElement) =>
        !pseudoElement && (element.LocalName.ToLowerInvariant() is "img" or "svg" or "iframe"
            or "input" or "textarea" or "select" or "button" or "progress" or "meter" or "math");
}
