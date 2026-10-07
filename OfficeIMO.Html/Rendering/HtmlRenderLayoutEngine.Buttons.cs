using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Unlike input button values, HTML button descendants have their own CSS
    // boxes, fonts and visibility. Keep that content in the normal layout owner.
    private static bool UsesButtonChildLayout(IElement element) =>
        string.Equals(element.LocalName, "button", StringComparison.OrdinalIgnoreCase)
        && element.ChildElementCount > 0;

    private HtmlRenderBoxStyle PrepareButtonChildStyle(IElement element, HtmlRenderBoxStyle authoredStyle) {
        if (!UsesButtonChildLayout(element) || authoredStyle.Display == "none") return authoredStyle;

        HtmlRenderBoxStyle style = CreateFormControlStyle(element, authoredStyle);
        if (!style.DisplayWasSpecified) style.Display = "inline-block";
        if (_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)) {
            bool Declared(string property) => computed.IsSpecifiedValue(property) || computed.IsInheritedValue(property);
            if (Declared("background") || Declared("background-color")) style.BackgroundColor = authoredStyle.BackgroundColor;
            bool padding = Declared("padding");
            if (padding || Declared("padding-left")) style.PaddingLeft = authoredStyle.PaddingLeft;
            if (padding || Declared("padding-right")) style.PaddingRight = authoredStyle.PaddingRight;
            if (padding || Declared("padding-top")) style.PaddingTop = authoredStyle.PaddingTop;
            if (padding || Declared("padding-bottom")) style.PaddingBottom = authoredStyle.PaddingBottom;
            if (!computed.IsSpecifiedValue("text-align")) style.Alignment = OfficeTextAlignment.Center;
        }
        return style;
    }

    // Native buttons center normal-flow content in the unused block-axis space.
    // Explicit flex/grid layout has its own alignment owner and returns earlier.
    private static double ResolveButtonChildContentOffset(
        IElement element, HtmlRenderBoxStyle style, double boxHeight, double contentHeight) =>
        UsesButtonChildLayout(element) && !IsVerticalWritingMode(style.WritingMode)
            ? Math.Max(0D, (boxHeight - style.VerticalInsets - contentHeight) / 2D)
            : 0D;
}
