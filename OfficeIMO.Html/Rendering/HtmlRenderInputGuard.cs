using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlRenderInputGuard {
    internal static void ValidateSource(string html, HtmlRenderOptions options) {
        if (html.Length <= options.MaxInputCharacters) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
            "HTML source length exceeded the configured render limit.",
            nameof(HtmlRenderOptions.MaxInputCharacters),
            html.Length,
            options.MaxInputCharacters);
    }

    internal static void ValidateDocument(IHtmlDocument document, HtmlRenderOptions options, CancellationToken cancellationToken) {
        var limits = new HtmlDomLimitTracker(options.MaxHtmlNodes, null);
        var pending = new Stack<INode>();
        PushChildren(document, pending);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            INode node = pending.Pop();
            limits.RecordNode();
            PushChildren(node, pending);
        }
    }

    private static void PushChildren(INode node, Stack<INode> pending) {
        for (int index = node.ChildNodes.Length - 1; index >= 0; index--) {
            pending.Push(node.ChildNodes[index]);
        }
    }
}
