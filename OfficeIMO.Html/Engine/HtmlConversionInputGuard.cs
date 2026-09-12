using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Applies shared source and DOM limits before expensive conversion analysis begins.</summary>
internal static class HtmlConversionInputGuard {
    internal const int MaxSrcDocDepth = 8;

    internal static void ValidateSource(string html, HtmlConversionLimits limits) {
        if (!limits.MaxInputCharacters.HasValue || html.Length <= limits.MaxInputCharacters.Value) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
            "HTML source length exceeded the configured conversion limit.",
            nameof(HtmlConversionLimits.MaxInputCharacters),
            html.Length,
            limits.MaxInputCharacters.Value);
    }

    internal static void ValidateDocument(IDocument document, HtmlConversionLimits limits, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlDomLimitTracker? tracker = HtmlDomLimitTracker.Create(limits.MaxHtmlNodes, limits.MaxHtmlDepth);
        var cssBudget = new HtmlCssByteBudget(limits);

        // Retain one enumerator per ancestor, rather than buffering every sibling before checking its budget.
        var pending = new Stack<(IEnumerator<INode> Nodes, int ParentDepth, int SrcDocDepth)>();
        pending.Push((document.ChildNodes.GetEnumerator(), 0, 0));
        try {
            while (pending.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                var level = pending.Peek();
                if (!level.Nodes.MoveNext()) { level.Nodes.Dispose(); pending.Pop(); continue; }
                INode node = level.Nodes.Current;
                int depth = level.ParentDepth + (node is IElement ? 1 : 0);
                if (node is IElement element) {
                    tracker?.RecordElementStart(depth);
                    ValidateSemanticAttributes(element, limits.MaxSemanticMetadataCharacters, cancellationToken);
                    if (string.Equals(element.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                        cssBudget.ReserveOrThrow(element.TextContent ?? string.Empty);
                    }
                    if (level.SrcDocDepth < MaxSrcDocDepth) {
                        string? source = element.GetAttribute("srcdoc");
                        if (!string.IsNullOrWhiteSpace(source)) {
                            ValidateSource(source!, limits);
                            IHtmlDocument nested = HtmlDocumentParser.ParseDocument(source!, cancellationToken);
                            pending.Push((nested.ChildNodes.GetEnumerator(), depth, level.SrcDocDepth + 1));
                        }
                    }
                    if (element is IHtmlTemplateElement template) {
                        pending.Push((((IEnumerable<INode>)new[] { template.Content }).GetEnumerator(), depth, level.SrcDocDepth));
                    }
                } else {
                    tracker?.RecordNode();
                }
                pending.Push((node.ChildNodes.GetEnumerator(), depth, level.SrcDocDepth));
            }
        } finally {
            while (pending.Count != 0) pending.Pop().Nodes.Dispose();
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void ValidateSemanticAttributes(IElement element, int? maximumCharacters, CancellationToken cancellationToken) {
        if (!maximumCharacters.HasValue) return;
        foreach (IAttr attribute in element.Attributes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!attribute.Name.StartsWith("data-officeimo-", StringComparison.OrdinalIgnoreCase)
                || attribute.Value.Length <= maximumCharacters.Value) {
                continue;
            }

            throw new HtmlDomLimitException(
                HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "OfficeIMO semantic metadata exceeded the configured conversion limit.",
                nameof(HtmlConversionLimits.MaxSemanticMetadataCharacters),
                attribute.Value.Length,
                maximumCharacters.Value);
        }
    }

}
