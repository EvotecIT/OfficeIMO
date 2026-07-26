using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private string? ResolveRunningStringContentText(
        IElement element,
        HtmlRenderBoxStyle elementStyle,
        int maximumCharacters,
        Action<long> chargeOperations) {
        var result = new StringBuilder(Math.Min(Math.Max(0, maximumCharacters), 256));
        var pending = new Stack<(INode Node, HtmlRenderBoxStyle ParentStyle)>();
        for (int index = element.ChildNodes.Length - 1; index >= 0; index--) {
            pending.Push((element.ChildNodes[index], elementStyle));
        }

        bool whitespace = false;
        int pendingCharacterCharge = 0;
        double containingWidth = Math.Max(
            1D,
            (_options.Mode == HtmlRenderMode.Paged ? _options.PageWidth : _options.ViewportWidth)
            - _options.Margins.Left
            - _options.Margins.Right);
        while (pending.Count > 0) {
            (INode node, HtmlRenderBoxStyle parentStyle) = pending.Pop();
            chargeOperations(1L);
            if (node is IText textNode) {
                pendingCharacterCharge += textNode.Data.Length;
                if (pendingCharacterCharge >= 256) {
                    chargeOperations(pendingCharacterCharge);
                    pendingCharacterCharge = 0;
                }
                if (!parentStyle.PaintVisible) continue;

                string transformed = ApplyTextTransform(textNode.Data, parentStyle.TextTransform);
                foreach (char current in transformed) {
                    if (char.IsWhiteSpace(current)) {
                        whitespace = result.Length > 0;
                        continue;
                    }
                    int required = whitespace ? 2 : 1;
                    if (result.Length > maximumCharacters - required) {
                        if (pendingCharacterCharge > 0) chargeOperations(pendingCharacterCharge);
                        return null;
                    }
                    if (whitespace) result.Append(' ');
                    result.Append(current);
                    whitespace = false;
                }
                continue;
            }

            if (node is not IElement child || ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle childStyle = _layoutStyles.TryGetValue(child, out HtmlRenderBoxStyle? cachedStyle)
                ? cachedStyle
                : _styleResolver.Resolve(child, containingWidth, parentStyle);
            if (childStyle.Display == "none") continue;
            for (int index = child.ChildNodes.Length - 1; index >= 0; index--) {
                pending.Push((child.ChildNodes[index], childStyle));
            }
        }

        if (pendingCharacterCharge > 0) chargeOperations(pendingCharacterCharge);
        return result.ToString();
    }
}
