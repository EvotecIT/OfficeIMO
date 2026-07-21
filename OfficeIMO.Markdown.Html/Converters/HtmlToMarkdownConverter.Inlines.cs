using AngleSharp.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
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
                AppendNormalizedText(sequence, text.Data, trimEnd, context);
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
        if (IsPassThroughTag(element, context)) {
            sequence.AddRaw(new HtmlRawInline(NormalizeRawElement(element, context)));
            return;
        }

        if (TryConvertConfiguredInlineElementConverters(sequence, element, context)) {
            return;
        }

        if (context != null && element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) {
            if (context.Footnotes.IsBacklink(element)) return;
            if (context.Footnotes.TryGetReferenceLabel(element, out string footnoteLabel)) {
                sequence.FootnoteRef(footnoteLabel);
                return;
            }
        }
        if (context != null && context.Footnotes.TryGetWrappedReferenceLabel(element, out string wrappedFootnoteLabel)) {
            sequence.FootnoteRef(wrappedFootnoteLabel);
            return;
        }

        string tag = GetEffectiveTagName(element, context);
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
                AppendInlineElementChildren(sequence, element, context);
                return;
            default:
                AppendUnknownInlineElement(sequence, element, context);
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
                if (IsPassThroughTag(element, context)) {
                    return NormalizeRawElement(element, context);
                }
                return ConvertInlineElementToMarkdown(element, context);
            default:
                return string.Empty;
        }
    }

    private static string ConvertInlineElementToMarkdown(IElement element, ConversionContext? context) {
        if (context != null && element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) {
            if (context.Footnotes.IsBacklink(element)) return string.Empty;
            if (context.Footnotes.TryGetReferenceLabel(element, out string footnoteLabel)) {
                return "[^" + footnoteLabel + "]";
            }
        }
        if (context != null && context.Footnotes.TryGetWrappedReferenceLabel(element, out string wrappedFootnoteLabel)) {
            return "[^" + wrappedFootnoteLabel + "]";
        }

        string tag = GetEffectiveTagName(element, context);
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
                return ConvertUnknownInlineElementToMarkdown(element, context);
        }
    }

    private static string ConvertHtmlWrappedInlineElementToMarkdown(IElement element, ConversionContext? context, string tagName) {
        return "<" + tagName + ">" + ConvertInlineNodesToMarkdown(element.ChildNodes, context) + "</" + tagName + ">";
    }

    private static string ConvertInlineInputToMarkdown(IElement element, ConversionContext? context) {
        string? type = element.GetAttribute("type");
        if (string.Equals(type, "checkbox", StringComparison.OrdinalIgnoreCase)) {
            if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
                return NormalizeRawElement(element, context);
            }

            return element.HasAttribute("checked") ? "[x]" : "[ ]";
        }

        if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
            return NormalizeRawElement(element, context);
        }

        return string.Empty;
    }

    private static string ConvertInlineAnchorToMarkdown(IElement element, ConversionContext? context) {
        string href = context == null
            ? element.GetAttribute("href") ?? string.Empty
            : ResolveUrl(element.GetAttribute("href"), context);
        string label = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
        if (label.Length == 0) {
            label = EscapeInlineText(HtmlAccessibilitySemantics.GetAccessibleName(element, includeTextFallback: true));
            if (label.Length == 0) return string.Empty;
        }
        if (href.Length == 0) {
            return label;
        }

        if (context?.Options.SmartHref == true && TryConvertSmartHrefToPlainText(label, href, out string plainText)) {
            return EscapeInlineText(plainText);
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
        if (label.Nodes.Count == 0) {
            string accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(element, includeTextFallback: true);
            if (accessibleName.Length == 0) return;
            label.Text(accessibleName);
        }

        if (href.Length == 0) {
            foreach (var node in label.Nodes) {
                sequence.AddRaw(node);
            }
            return;
        }

        if (context?.Options.SmartHref == true
            && TryConvertSmartHrefToPlainText(ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim(), href, out string plainText)) {
            sequence.Text(plainText);
            return;
        }

        sequence.AddRaw(new LinkInline(
            label,
            href,
            element.GetAttribute("title"),
            element.GetAttribute("target"),
            element.GetAttribute("rel")));
    }

    private static bool TryConvertSmartHrefToPlainText(string label, string href, out string plainText) {
        plainText = string.Empty;
        if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(href)) {
            return false;
        }

        string normalizedLabel = NormalizeSmartHrefText(label);
        string normalizedHref = href.Trim();
        if (normalizedLabel.Length == 0 || normalizedHref.Length == 0) {
            return false;
        }

        if (string.Equals(normalizedLabel, normalizedHref, StringComparison.OrdinalIgnoreCase)) {
            plainText = normalizedHref;
            return true;
        }

        if (normalizedHref.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
            && string.Equals(normalizedLabel, normalizedHref.Substring("mailto:".Length), StringComparison.OrdinalIgnoreCase)) {
            plainText = normalizedLabel;
            return true;
        }

        if (normalizedHref.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            && string.Equals(normalizedLabel, normalizedHref.Substring("http://".Length), StringComparison.OrdinalIgnoreCase)) {
            plainText = normalizedHref;
            return true;
        }

        if (normalizedHref.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            && string.Equals(normalizedLabel, normalizedHref.Substring("https://".Length), StringComparison.OrdinalIgnoreCase)) {
            plainText = normalizedHref;
            return true;
        }

        return false;
    }

    private static string NormalizeSmartHrefText(string value) {
        return value
            .Replace("\\.", ".")
            .Replace("\\-", "-")
            .Replace("\\_", "_")
            .Trim();
    }

    private static string ConvertInlineImageToMarkdown(IElement element, ConversionContext? context) {
        var metadata = ResolveInlineImageMetadata(element, context);
        if (metadata.src.Length == 0) {
            return string.Empty;
        }

        string escapedAlt = EscapeInlineText(metadata.alt);
        string normalizedTitle = metadata.title ?? string.Empty;
        string titlePart = normalizedTitle.Length == 0
            ? string.Empty
            : " \"" + normalizedTitle.Replace("\"", "\\\"") + "\"";
        return "![" + escapedAlt + "](" + EscapeLinkTarget(metadata.src) + titlePart + ")";
    }

    private static void AppendInlineImage(InlineSequence sequence, IElement element, ConversionContext? context) {
        var metadata = ResolveInlineImageMetadata(element, context);
        if (metadata.src.Length == 0) {
            return;
        }

        sequence.AddRaw(new ImageInline(
            metadata.alt ?? string.Empty,
            metadata.src,
            metadata.title,
            metadata.plainAlt));
    }

    private static (string src, string? alt, string? title, string? plainAlt) ResolveInlineImageMetadata(IElement element, ConversionContext? context) {
        if (context != null && TryCreateImageBlock(element, context, out var image)) {
            return (image.Path, image.Alt, image.Title, image.PlainAlt);
        }

        if (context != null
            && context.Options.Base64Images != HtmlBase64ImageHandling.Include
            && ResolveImageSourceCandidates(element, context).Any(IsBase64ImageDataUri)) {
            string? accessibleName = GetAccessibleImageName(element);
            return (string.Empty, accessibleName, element.GetAttribute("title"), accessibleName);
        }

        string src = context == null
            ? element.GetAttribute("src") ?? string.Empty
            : ResolveResourceUrl(element.GetAttribute("src"), context);
        string? alt = GetAccessibleImageName(element);
        return (src, alt, element.GetAttribute("title"), alt);
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

    private static void AppendNormalizedText(InlineSequence sequence, string? text, bool trimEnd, ConversionContext? context) {
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

        if (context?.Options.EscapeMarkdownLineStarts == true) {
            if (sequence.Nodes.Count > 0 && sequence.Nodes[sequence.Nodes.Count - 1] is LineStartEscapedTextRun previous) {
                var nodes = sequence.Nodes.Take(sequence.Nodes.Count - 1).ToList();
                nodes.Add(new LineStartEscapedTextRun(previous.Text + normalized));
                sequence.ReplaceItems(nodes);
                return;
            }

            sequence.AddRaw(new LineStartEscapedTextRun(normalized));
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
                sequence.AddRaw(new HtmlRawInline(NormalizeRawElement(element, context)));
                return;
            }

            sequence.Text(element.HasAttribute("checked") ? "[x]" : "[ ]");
            return;
        }

        if (context != null && context.Options.PreserveUnsupportedInlineHtml) {
            sequence.AddRaw(new HtmlRawInline(NormalizeRawElement(element, context)));
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
