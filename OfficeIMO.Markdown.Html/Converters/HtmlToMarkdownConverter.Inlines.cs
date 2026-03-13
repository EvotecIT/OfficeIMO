using AngleSharp.Dom;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static string ConvertInlineNodesToMarkdown(IEnumerable<INode> nodes, ConversionContext? context) {
        var sb = new StringBuilder();
        foreach (var node in nodes) {
            sb.Append(ConvertInlineNodeToMarkdown(node, context));
        }
        return sb.ToString();
    }

    private static string ConvertInlineNodeToMarkdown(INode node, ConversionContext? context) {
        switch (node) {
            case IComment:
                return string.Empty;
            case IText text:
                return EscapeInlineText(text.Data);
            case IElement element:
                if (context != null && ShouldIgnoreElement(element, context)) {
                    return string.Empty;
                }
                return ConvertInlineElementToMarkdown(element, context);
            default:
                return string.Empty;
        }
    }

    private static string ConvertInlineElementToMarkdown(IElement element, ConversionContext? context) {
        string tag = element.TagName;
        switch (tag) {
            case "BR":
                return "  \n";
            case "STRONG":
            case "B":
                return "**" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "**";
            case "EM":
            case "I":
                return "*" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "*";
            case "S":
            case "STRIKE":
            case "DEL":
                return "~~" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "~~";
            case "MARK":
                return "==" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "==";
            case "U":
                return "<u>" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "</u>";
            case "CODE":
                return WrapCode(element.TextContent ?? string.Empty);
            case "A":
                return ConvertInlineAnchorToMarkdown(element, context);
            case "IMG":
                return ConvertInlineImageToMarkdown(element, context);
            case "SPAN":
            case "SMALL":
            case "BIG":
            case "ABBR":
            case "CITE":
            case "TIME":
            case "KBD":
            case "SAMP":
            case "VAR":
            case "LABEL":
                return ConvertInlineNodesToMarkdown(element.ChildNodes, context);
            default:
                if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
                    return element.OuterHtml;
                }
                return ConvertInlineNodesToMarkdown(element.ChildNodes, context);
        }
    }

    private static string ConvertInlineAnchorToMarkdown(IElement element, ConversionContext? context) {
        string href = context == null
            ? element.GetAttribute("href") ?? string.Empty
            : ResolveUrl(element.GetAttribute("href"), context);
        string label = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
        if (label.Length == 0) {
            label = EscapeInlineText(href);
        }
        if (href.Length == 0) {
            return label;
        }

        string title = element.GetAttribute("title") ?? string.Empty;
        string titlePart = title.Length == 0
            ? string.Empty
            : " \"" + title.Replace("\"", "\\\"") + "\"";
        return "[" + label + "](" + EscapeLinkTarget(href) + titlePart + ")";
    }

    private static string ConvertInlineImageToMarkdown(IElement element, ConversionContext? context) {
        string src = context == null
            ? element.GetAttribute("src") ?? string.Empty
            : ResolveUrl(element.GetAttribute("src"), context);
        if (src.Length == 0) {
            return string.Empty;
        }

        string alt = EscapeInlineText(element.GetAttribute("alt"));
        string title = element.GetAttribute("title") ?? string.Empty;
        string titlePart = title.Length == 0
            ? string.Empty
            : " \"" + title.Replace("\"", "\\\"") + "\"";
        return "![" + alt + "](" + EscapeLinkTarget(src) + titlePart + ")";
    }
}
