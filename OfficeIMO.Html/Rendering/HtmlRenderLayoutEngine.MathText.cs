using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private string ResolveMathTokenPaint(IElement token, IElement root, HtmlRenderBoxStyle rootStyle, double containingWidth) {
        // Resolve each token through the same cascade as HTML text. The MathML model
        // remains logical: its native source supplement never contains presentation glyph aliases.
        var styles = new Dictionary<IElement, HtmlRenderBoxStyle> { [root] = rootStyle };
        HtmlRenderBoxStyle Resolve(IElement owner) {
            if (styles.TryGetValue(owner, out HtmlRenderBoxStyle? found)) return found;
            HtmlRenderBoxStyle parent = owner.ParentElement == null ? rootStyle : Resolve(owner.ParentElement);
            return styles[owner] = _styleResolver.Resolve(owner, containingWidth, parent);
        }
        string Paint(INode node, HtmlRenderBoxStyle parent) {
            CheckCancellation();
            if (node.NodeType == NodeType.Text) return ApplyTextTransform(node.TextContent, parent);
            if (node is not IElement child) return string.Empty;
            HtmlRenderBoxStyle style = Resolve(child);
            return string.Concat(child.ChildNodes.Select(item => Paint(item, style)));
        }
        return Paint(token, rootStyle);
    }

    private sealed class MathTokenIdentityComparer : IEqualityComparer<OfficeMathExpression> {
        internal static readonly MathTokenIdentityComparer Instance = new();
        public bool Equals(OfficeMathExpression? x, OfficeMathExpression? y) => ReferenceEquals(x, y);
        public int GetHashCode(OfficeMathExpression value) => RuntimeHelpers.GetHashCode(value);
    }
}
