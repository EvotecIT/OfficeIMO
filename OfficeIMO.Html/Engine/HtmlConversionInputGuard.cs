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

    internal static void ValidateDocument(IDocument document, HtmlConversionLimits limits) {
        HtmlDomLimitTracker? tracker = HtmlDomLimitTracker.Create(limits.MaxHtmlNodes, limits.MaxHtmlDepth);
        long totalCssBytes = 0L;

        var pending = new Stack<(INode Node, int Depth, int SrcDocDepth)>();
        for (int index = document.ChildNodes.Length - 1; index >= 0; index--) {
            pending.Push((document.ChildNodes[index], 1, 0));
        }

        while (pending.Count > 0) {
            (INode node, int depth, int srcDocDepth) = pending.Pop();
            if (node is IElement element) {
                tracker?.RecordElementStart(depth);
                ValidateSemanticAttributes(element, limits.MaxSemanticMetadataCharacters);
                if (string.Equals(element.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                    ValidateStylesheet(element.TextContent ?? string.Empty, limits, ref totalCssBytes);
                }
                if (srcDocDepth < MaxSrcDocDepth) {
                    PushSrcDoc(element, depth, srcDocDepth, limits, pending);
                }
            } else {
                tracker?.RecordNode();
            }
            for (int index = node.ChildNodes.Length - 1; index >= 0; index--) {
                pending.Push((node.ChildNodes[index], depth + 1, srcDocDepth));
            }
        }
    }

    private static void PushSrcDoc(
        IElement element,
        int parentDepth,
        int srcDocDepth,
        HtmlConversionLimits limits,
        Stack<(INode Node, int Depth, int SrcDocDepth)> pending) {
        string? source = element.GetAttribute("srcdoc");
        if (string.IsNullOrWhiteSpace(source)) return;

        ValidateSource(source!, limits);
        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(source!);
        for (int index = nested.ChildNodes.Length - 1; index >= 0; index--) {
            pending.Push((nested.ChildNodes[index], parentDepth + 1, srcDocDepth + 1));
        }
    }

    private static void ValidateSemanticAttributes(IElement element, int? maximumCharacters) {
        if (!maximumCharacters.HasValue) return;
        foreach (IAttr attribute in element.Attributes) {
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

    private static void ValidateStylesheet(string css, HtmlConversionLimits limits, ref long totalBytes) {
        if (!limits.MaxCssBytes.HasValue && !limits.MaxTotalCssBytes.HasValue) return;
        long bytes = Encoding.UTF8.GetByteCount(css);
        if (limits.MaxCssBytes.HasValue && bytes > limits.MaxCssBytes.Value) {
            throw CreateCssLimitException(
                HtmlConversionDiagnosticCodes.CssSizeLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssBytes),
                bytes,
                limits.MaxCssBytes.Value);
        }

        totalBytes += bytes;
        if (limits.MaxTotalCssBytes.HasValue && totalBytes > limits.MaxTotalCssBytes.Value) {
            throw CreateCssLimitException(
                HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded,
                nameof(HtmlConversionLimits.MaxTotalCssBytes),
                totalBytes,
                limits.MaxTotalCssBytes.Value);
        }
    }

    private static HtmlDomLimitException CreateCssLimitException(string code, string source, long actual, long limit) =>
        new HtmlDomLimitException(code, "Embedded CSS exceeded the configured conversion limit.", source, actual, limit);
}
