using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<INode, (int Start, int End)> _semanticSourceOrders = new();

    // Text and generated children share the same source coordinate as element
    // containers. Entry/exit positions put ::after beyond every authored child.
    private void InitializeSemanticSourceOrder(INode document) {
        var pending = new Stack<(INode Node, bool Exiting)>();
        pending.Push((document, false));
        int order = 0;
        while (pending.Count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            var entry = pending.Pop();
            if (entry.Exiting) {
                var position = _semanticSourceOrders[entry.Node];
                _semanticSourceOrders[entry.Node] = (position.Start, order++);
                continue;
            }
            _semanticSourceOrders[entry.Node] = (order++, 0);
            pending.Push((entry.Node, true));
            for (INode? child = entry.Node.LastChild; child != null; child = child.PreviousSibling)
                pending.Push((child, false));
        }
    }

    private int? GetSemanticSourceOrder(INode node) =>
        _semanticSourceOrders.TryGetValue(node, out var position) ? position.Start : null;

    private int? GetGeneratedSemanticSourceOrder(IElement element, HtmlPseudoElementKind kind) =>
        _semanticSourceOrders.TryGetValue(element, out var position)
            ? kind == HtmlPseudoElementKind.After ? position.End : position.Start
            : null;
}
