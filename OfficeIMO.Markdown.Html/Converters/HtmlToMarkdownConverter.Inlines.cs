using AngleSharp.Dom;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    internal static InlineSequence ConvertInlineNodesToInlineSequence(IEnumerable<INode> nodes, ConversionContext? context) {
        var sequence = new InlineSequence { AutoSpacing = false };
        var materializedNodes = nodes as IList<INode> ?? nodes.ToList();
        for (int i = 0; i < materializedNodes.Count; i++) {
            bool trimEnd = NextVisibleInlineNodeIsBoundary(materializedNodes, i + 1, context);
            AppendInlineNode(sequence, materializedNodes[i], context, trimEnd);
        }
        return sequence;
    }

    private static string ConvertInlineNodesToMarkdown(IEnumerable<INode> nodes, ConversionContext? context) {
        var sb = new StringBuilder();
        foreach (var node in nodes) {
            sb.Append(ConvertInlineNodeToMarkdown(node, context));
        }
        return sb.ToString();
    }

    private static void AppendInlineNode(InlineSequence sequence, INode node, ConversionContext? context, bool trimEnd) {
        switch (node) {
            case null:
            case IComment:
                return;
            case IText text:
                AppendNormalizedText(sequence, text.Data, trimEnd);
                return;
            case IElement element:
                if (context != null && ShouldIgnoreElement(element, context)) {
                    return;
                }
                AppendInlineElement(sequence, element, context);
                return;
        }
    }

    private static void AppendInlineElement(InlineSequence sequence, IElement element, ConversionContext? context) {
        if (TryConvertConfiguredInlineElementConverters(sequence, element, context)) {
            return;
        }

        string tag = element.TagName;
        switch (tag) {
            case "BR":
                sequence.HardBreak();
                return;
            case "STRONG":
            case "B":
                AppendWrappedInlineSequence(sequence, ConvertInlineNodesToInlineSequence(element.ChildNodes, context), static child => new BoldSequenceInline(child));
                return;
            case "EM":
            case "I":
                AppendWrappedInlineSequence(sequence, ConvertInlineNodesToInlineSequence(element.ChildNodes, context), static child => new ItalicSequenceInline(child));
                return;
            case "S":
            case "STRIKE":
            case "DEL":
                AppendWrappedInlineSequence(sequence, ConvertInlineNodesToInlineSequence(element.ChildNodes, context), static child => new StrikethroughSequenceInline(child));
                return;
            case "MARK":
                AppendWrappedInlineSequence(sequence, ConvertInlineNodesToInlineSequence(element.ChildNodes, context), static child => new HighlightSequenceInline(child));
                return;
            case "U":
            case "SUP":
            case "SUB":
            case "INS":
            case "Q":
                AppendWrappedInlineSequence(sequence, ConvertInlineNodesToInlineSequence(element.ChildNodes, context), child => new HtmlTagSequenceInline(tag.ToLowerInvariant(), child));
                return;
            case "CODE":
                sequence.AddRaw(new CodeSpanInline(element.TextContent ?? string.Empty));
                return;
            case "A":
                AppendInlineAnchor(sequence, element, context);
                return;
            case "IMG":
                AppendInlineImage(sequence, element, context);
                return;
            case "INPUT":
                AppendInlineInput(sequence, element, context);
                return;
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
                var childNodes = element.ChildNodes as IList<INode> ?? element.ChildNodes.ToList();
                for (int i = 0; i < childNodes.Count; i++) {
                    bool trimEnd = NextVisibleInlineNodeIsBoundary(childNodes, i + 1, context);
                    AppendInlineNode(sequence, childNodes[i], context, trimEnd);
                }
                return;
            default:
                if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
                    sequence.AddRaw(new HtmlRawInline(element.OuterHtml));
                    return;
                }

                var fallbackNodes = element.ChildNodes as IList<INode> ?? element.ChildNodes.ToList();
                for (int i = 0; i < fallbackNodes.Count; i++) {
                    bool trimEnd = NextVisibleInlineNodeIsBoundary(fallbackNodes, i + 1, context);
                    AppendInlineNode(sequence, fallbackNodes[i], context, trimEnd);
                }
                return;
        }
    }

    private static bool TryConvertConfiguredInlineElementConverters(InlineSequence sequence, IElement element, ConversionContext? context) {
        if (sequence == null || element == null || context?.Options?.InlineElementConverters == null || context.Options.InlineElementConverters.Count == 0) {
            return false;
        }

        var conversionContext = new HtmlInlineElementConversionContext(element, context.Options, context);
        for (int i = 0; i < context.Options.InlineElementConverters.Count; i++) {
            var converter = context.Options.InlineElementConverters[i];
            if (converter == null) {
                continue;
            }

            var converted = converter.Convert(conversionContext);
            if (converted == null) {
                continue;
            }

            for (int j = 0; j < converted.Count; j++) {
                var inline = converted[j];
                if (inline != null) {
                    sequence.AddRaw(inline);
                }
            }

            return true;
        }

        return false;
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
                return ConvertStrongInlineElementToMarkdown(element, context);
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
                return ConvertHtmlWrappedInlineElementToMarkdown(element, context, "u");
            case "SUP":
                return ConvertHtmlWrappedInlineElementToMarkdown(element, context, "sup");
            case "SUB":
                return ConvertHtmlWrappedInlineElementToMarkdown(element, context, "sub");
            case "INS":
                return ConvertHtmlWrappedInlineElementToMarkdown(element, context, "ins");
            case "Q":
                return ConvertHtmlWrappedInlineElementToMarkdown(element, context, "q");
            case "CODE":
                return WrapCode(element.TextContent ?? string.Empty);
            case "A":
                return ConvertInlineAnchorToMarkdown(element, context);
            case "IMG":
                return ConvertInlineImageToMarkdown(element, context);
            case "INPUT":
                return ConvertInlineInputToMarkdown(element, context);
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

    private static string ConvertHtmlWrappedInlineElementToMarkdown(IElement element, ConversionContext? context, string tagName) {
        return "<" + tagName + ">" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "</" + tagName + ">";
    }

    private static string ConvertInlineInputToMarkdown(IElement element, ConversionContext? context) {
        string? type = element.GetAttribute("type");
        if (string.Equals(type, "checkbox", StringComparison.OrdinalIgnoreCase)) {
            if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
                return element.OuterHtml;
            }

            return element.HasAttribute("checked") ? "[x]" : "[ ]";
        }

        if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
            return element.OuterHtml;
        }

        return string.Empty;
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

    private static void AppendInlineAnchor(InlineSequence sequence, IElement element, ConversionContext? context) {
        string href = context == null
            ? element.GetAttribute("href") ?? string.Empty
            : ResolveUrl(element.GetAttribute("href"), context);
        var label = ConvertInlineNodesToInlineSequence(element.ChildNodes, context);
        if (label.Nodes.Count == 0 && href.Length > 0) {
            label.Text(href);
        }

        if (href.Length == 0) {
            foreach (var node in label.Nodes) {
                sequence.AddRaw(node);
            }
            return;
        }

        sequence.AddRaw(new LinkInline(
            label,
            href,
            element.GetAttribute("title"),
            element.GetAttribute("target"),
            element.GetAttribute("rel")));
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

    private static void AppendInlineImage(InlineSequence sequence, IElement element, ConversionContext? context) {
        string src = context == null
            ? element.GetAttribute("src") ?? string.Empty
            : ResolveUrl(element.GetAttribute("src"), context);
        if (src.Length == 0) {
            return;
        }

        sequence.AddRaw(new ImageInline(
            element.GetAttribute("alt") ?? string.Empty,
            src,
            element.GetAttribute("title")));
    }

    private static void AppendWrappedInlineSequence(
        InlineSequence target,
        InlineSequence childSequence,
        Func<InlineSequence, IMarkdownInline> factory) {
        if (target == null || childSequence == null || factory == null) {
            return;
        }

        if (childSequence.Nodes.Count == 0) {
            return;
        }

        target.AddRaw(factory(childSequence));
    }

    private static void AppendNormalizedText(InlineSequence sequence, string? text, bool trimEnd) {
        if (sequence == null || string.IsNullOrEmpty(text)) {
            return;
        }

        string normalized = CollapseHtmlInlineWhitespace(text!);
        if (normalized.Length == 0) {
            return;
        }

        bool trimLeading = sequence.Nodes.Count == 0 || sequence.Nodes[sequence.Nodes.Count - 1] is HardBreakInline;
        if (trimLeading) {
            normalized = normalized.TrimStart(' ');
        }

        if (trimEnd) {
            normalized = normalized.TrimEnd(' ');
        }

        if (normalized.Length == 0) {
            return;
        }

        sequence.Text(normalized);
    }

    internal static string CollapseHtmlInlineWhitespace(string text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        var sb = new StringBuilder(text.Length);
        bool previousWasWhitespace = false;
        foreach (char ch in text) {
            if (char.IsWhiteSpace(ch)) {
                if (previousWasWhitespace) {
                    continue;
                }

                sb.Append(' ');
                previousWasWhitespace = true;
                continue;
            }

            sb.Append(ch);
            previousWasWhitespace = false;
        }

        return sb.ToString();
    }

    private static bool NextVisibleInlineNodeIsBoundary(IList<INode> nodes, int startIndex, ConversionContext? context) {
        for (int i = startIndex; i < nodes.Count; i++) {
            var node = nodes[i];
            switch (node) {
                case null:
                case IComment:
                    continue;
                case IText text when string.IsNullOrWhiteSpace(text.Data):
                    continue;
                case IElement element when context != null && ShouldIgnoreElement(element, context):
                    continue;
                case IElement element when element.TagName.Equals("BR", StringComparison.OrdinalIgnoreCase):
                    return true;
                default:
                    return false;
            }
        }

        return true;
    }

    private static string ConvertStrongInlineElementToMarkdown(IElement element, ConversionContext? context) {
        string content = TrimInlineMarkupBoundaryWhitespace(ConvertInlineNodesToMarkdown(element.ChildNodes, context));
        if (content.Length == 0) {
            return string.Empty;
        }

        return "**" + content + "**";
    }

    private static void AppendInlineInput(InlineSequence sequence, IElement element, ConversionContext? context) {
        string? type = element.GetAttribute("type");
        if (string.Equals(type, "checkbox", StringComparison.OrdinalIgnoreCase)) {
            if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
                sequence.AddRaw(new HtmlRawInline(element.OuterHtml));
                return;
            }

            sequence.Text(element.HasAttribute("checked") ? "[x]" : "[ ]");
            return;
        }

        if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
            sequence.AddRaw(new HtmlRawInline(element.OuterHtml));
        }
    }

    private static string TrimInlineMarkupBoundaryWhitespace(string value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        int start = 0;
        int end = value.Length - 1;

        while (start <= end && char.IsWhiteSpace(value[start])) {
            start++;
        }

        while (end >= start && char.IsWhiteSpace(value[end])) {
            end--;
        }

        return start > end
            ? string.Empty
            : value.Substring(start, end - start + 1);
    }
}
