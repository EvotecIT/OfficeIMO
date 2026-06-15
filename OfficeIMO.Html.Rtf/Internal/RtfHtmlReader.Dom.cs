using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private static void ReadDom(string html, RtfHtmlReadOptions options, RtfDocument document) {
        IHtmlDocument htmlDocument = HtmlDocumentParser.ParseDocument(html);
        Uri? effectiveBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(htmlDocument, options.BaseUri);
        var context = new ReadContext(document, options, effectiveBaseUri);
        HtmlDomLimitTracker? limits = HtmlDomLimitTracker.Create(options);
        TraverseNode(HtmlDocumentParser.GetConversionRoot(htmlDocument, useBodyContentsOnly: false), context, limits, depth: 0);
        context.TrimEmptyTrailingParagraph();
    }

    private static void TraverseNode(INode node, ReadContext context, HtmlDomLimitTracker? limits, int depth) {
        if (node is IText text) {
            limits?.RecordText();
            context.AppendText(text.Data);
            return;
        }

        if (node is IElement element) {
            string name = element.LocalName;
            bool closes = !IsVoidElement(name);
            limits?.RecordStart(depth + 1);
            context.Start(element);
            foreach (INode child in element.ChildNodes) {
                TraverseNode(child, context, limits, depth + 1);
            }

            if (closes) {
                context.End(name);
            }

            return;
        }

        foreach (INode child in node.ChildNodes) {
            TraverseNode(child, context, limits, depth);
        }
    }

    private static bool IsVoidElement(string name) {
        switch (name) {
            case "area":
            case "base":
            case "br":
            case "col":
            case "embed":
            case "hr":
            case "img":
            case "input":
            case "link":
            case "meta":
            case "param":
            case "source":
            case "track":
            case "wbr":
                return true;
            default:
                return false;
        }
    }

    private sealed class HtmlDomLimitTracker {
        private readonly RtfHtmlReadOptions _options;
        private int _nodes;

        private HtmlDomLimitTracker(RtfHtmlReadOptions options) {
            _options = options;
        }

        internal static HtmlDomLimitTracker? Create(RtfHtmlReadOptions? options) =>
            options != null && (options.MaxHtmlNodes.HasValue || options.MaxHtmlDepth.HasValue)
                ? new HtmlDomLimitTracker(options)
                : null;

        internal void RecordText() {
            RecordNode();
        }

        internal void RecordStart(int depth) {
            RecordNode();
            if (_options.MaxHtmlDepth.HasValue && depth > _options.MaxHtmlDepth.Value) {
                ThrowLimitExceeded("HtmlDepthLimitExceeded", "HTML nesting depth exceeded the configured conversion limit.", "MaxHtmlDepth", depth, _options.MaxHtmlDepth.Value);
            }
        }

        private void RecordNode() {
            _nodes++;
            if (_options.MaxHtmlNodes.HasValue && _nodes > _options.MaxHtmlNodes.Value) {
                ThrowLimitExceeded("HtmlNodeLimitExceeded", "HTML node count exceeded the configured conversion limit.", "MaxHtmlNodes", _nodes, _options.MaxHtmlNodes.Value);
            }
        }

        private void ThrowLimitExceeded(string code, string message, string source, long actual, long limit) {
            string detail = "Actual=" + actual + "; Limit=" + limit;
            var exception = new RtfHtmlConversionLimitException(code, message, source, actual, limit, detail);
            _options.AddDiagnostic(code, message, source, exception, RtfHtmlConversionDiagnosticSeverity.Error);
            throw exception;
        }
    }
}
