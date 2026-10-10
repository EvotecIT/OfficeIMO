using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool HasOrdinaryContainingHeightFormatting(HtmlRenderBoxStyle style) =>
        style.Display is "block" or "inline-block" or "inline" or "contents"
        && !IsVerticalWritingMode(style.WritingMode)
        && style.FloatSide == "none" && style.Position is "static" or "relative"
        && style.ColumnCount == null && style.ColumnWidth == null && style.ContainerType == "normal";

    /// <summary>
    /// Carries the enclosing block's content-height basis through ordinary non-replaced
    /// inline and contents wrappers. Their authored heights do not establish a new
    /// containing block; a missing enclosing height remains indefinite.
    /// </summary>
    private static HtmlRenderBoxStyle ForwardContainingHeightBasis(IElement element, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parent) {
        if (style.Display is not ("inline" or "contents")
            || !HasOrdinaryContainingHeightFormatting(style) || IsReplacedImageElement(element)
            || IsFormControlElement(element.LocalName) || element.LocalName is "svg" or "math" or "iframe") return style;
        HtmlRenderBoxStyle forwarded = style.Clone();
        forwarded.HasForwardedContainingHeight = true;
        forwarded.ForwardedContainingHeight = ResolveContainingBlockHeight(parent);
        return forwarded;
    }
}
