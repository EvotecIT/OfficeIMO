using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private static void ReadDom(IHtmlDocument htmlDocument, HtmlToRtfOptions options, RtfDocument document) {
        Uri? effectiveBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(htmlDocument, options.BaseUri);
        var context = new ReadContext(document, options, effectiveBaseUri);
        HtmlDomLimitTracker? limits = HtmlDomLimitTracker.Create(options.MaxHtmlNodes, options.MaxHtmlDepth);
        try {
            TraverseNode(HtmlDocumentParser.GetConversionRoot(htmlDocument, useBodyContentsOnly: false), context, limits, depth: 0);
        } catch (HtmlDomLimitException exception) {
            ThrowLimitExceeded(options, exception);
        }

        context.TrimEmptyTrailingParagraph();
    }

    private static void TraverseNode(INode node, ReadContext context, HtmlDomLimitTracker? limits, int depth) {
        if (node is IText text) {
            limits?.RecordNode();
            context.AppendText(text.Data);
            return;
        }

        if (node is IElement element) {
            string name = element.LocalName;
            bool closes = !HtmlDomElementFacts.IsVoidElement(name);
            limits?.RecordElementStart(depth + 1);
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

    private static void ThrowLimitExceeded(HtmlToRtfOptions options, HtmlDomLimitException exception) {
        var rtfException = new HtmlRtfConversionLimitException(
            exception.Code,
            exception.Message,
            exception.LimitSource,
            exception.Actual,
            exception.Limit,
            exception.Detail);
        options.AddDiagnostic(exception.Code, exception.Message, exception.LimitSource, rtfException, HtmlRtfConversionDiagnosticSeverity.Error, RtfConversionAction.Blocked);
        throw rtfException;
    }
}
