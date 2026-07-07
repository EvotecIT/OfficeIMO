using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static IEnumerable<IMarkdownBlock> ConvertParagraphElement(IElement element, ConversionContext context) {
        if (TryConvertParagraphAsStandaloneMediaBlock(element, context, out var mediaBlocks)) {
            return mediaBlocks;
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static bool TryConvertParagraphAsStandaloneMediaBlock(IElement element, ConversionContext context, out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (element == null || context == null) {
            return false;
        }

        IElement? onlyElement = null;
        foreach (var childNode in element.ChildNodes) {
            switch (childNode) {
                case IComment:
                    continue;
                case IText textNode when string.IsNullOrWhiteSpace(textNode.Data):
                    continue;
                case IElement childElement:
                    if (onlyElement != null) {
                        return false;
                    }

                    onlyElement = childElement;
                    break;
                default:
                    return false;
            }
        }

        if (onlyElement == null) {
            return false;
        }

        bool isStandaloneMedia = HasEffectiveTagName(onlyElement, context, "IMG")
                                 || HasEffectiveTagName(onlyElement, context, "PICTURE")
                                 || CanConvertAnchorToLinkedImageBlock(onlyElement, context);
        if (!isStandaloneMedia) {
            return false;
        }

        var converted = ConvertElementToBlocks(onlyElement, context).ToArray();
        if (converted.Length == 0) {
            return false;
        }

        blocks = converted;
        return true;
    }

    private static HeadingBlock ConvertHeadingElement(IElement element, int level, ConversionContext context) {
        return new HeadingBlock(level, NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context)));
    }

    private static IMarkdownBlock ConvertListElement(IElement element, bool ordered, ConversionContext context) {
        if (ordered) {
            var list = new OrderedListBlock();
            if (int.TryParse(element.GetAttribute("start"), out int start) && start > 0) {
                list.Start = start;
            }

            foreach (var item in element.Children.Where(child => HasEffectiveTagName(child, context, "LI"))) {
                list.Items.Add(ConvertListItem(item, context));
            }

            return list;
        }

        var unordered = new UnorderedListBlock();
        foreach (var item in element.Children.Where(child => HasEffectiveTagName(child, context, "LI"))) {
            unordered.Items.Add(ConvertListItem(item, context));
        }
        return unordered;
    }

    private static ListItem ConvertListItem(IElement element, ConversionContext context) {
        var filteredNodes = new List<INode>();
        bool isTask = false;
        bool isChecked = false;

        foreach (var child in element.ChildNodes) {
            if (child is IElement childElement
                && HasEffectiveTagName(childElement, context, "INPUT")
                && string.Equals(childElement.GetAttribute("type"), "checkbox", StringComparison.OrdinalIgnoreCase)) {
                isTask = true;
                isChecked = childElement.HasAttribute("checked");
                continue;
            }

            filteredNodes.Add(child);
        }

        var blocks = ConvertNodesToBlocks(filteredNodes, context);
        InlineSequence firstParagraph = new InlineSequence();
        int index = 0;
        if (blocks.Count > 0 && blocks[0] is ParagraphBlock first) {
            firstParagraph = first.Inlines;
            index = 1;
        }

        var item = isTask ? ListItem.TaskInlines(firstParagraph, isChecked) : new ListItem(firstParagraph);
        bool encounteredNonParagraphBlock = false;
        for (; index < blocks.Count; index++) {
            if (blocks[index] is ParagraphBlock paragraph) {
                if (!encounteredNonParagraphBlock) {
                    item.AdditionalParagraphs.Add(paragraph.Inlines);
                } else {
                    item.Children.Add(paragraph);
                }
            } else {
                encounteredNonParagraphBlock = true;
                item.Children.Add(blocks[index]);
            }
        }

        return item;
    }

    private static IMarkdownBlock ConvertBlockquoteElement(IElement element, ConversionContext context) {
        if (TryConvertCalloutElement(element, context, out var callout)) {
            return callout;
        }

        var quote = new QuoteBlock();
        foreach (var block in ConvertNodesToBlocks(element.ChildNodes, context)) {
            quote.Children.Add(block);
        }

        if (quote.Children.Count == 0) {
            string text = NormalizeBlockText(element.TextContent);
            if (text.Length > 0) {
                quote.Lines.Add(text);
            }
        }

        return quote;
    }

    private static bool TryConvertCalloutElement(IElement element, ConversionContext context, out CalloutBlock callout) {
        callout = null!;
        if (!element.ClassList.Contains("callout")) {
            return false;
        }

        string kind = "info";
        foreach (string token in element.ClassList) {
            if (string.IsNullOrWhiteSpace(token) || token.Equals("callout", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            kind = token.Trim().ToLowerInvariant();
            break;
        }

        var childBlocks = ConvertNodesToBlocks(element.ChildNodes, context);
        if (childBlocks.Count == 0) {
            callout = new CalloutBlock(kind, string.Empty, Array.Empty<IMarkdownBlock>());
            return true;
        }

        var blocks = new List<IMarkdownBlock>(childBlocks);
        var titleInlines = new InlineSequence();
        bool titleExplicit = IsCalloutTitleExplicit(element);
        if (!titleExplicit && blocks[0] is ParagraphBlock synthesizedTitleParagraph && IsStrongOnlyParagraph(synthesizedTitleParagraph)) {
            blocks.RemoveAt(0);
        } else if (blocks[0] is ParagraphBlock firstParagraph
            && TryExtractCalloutTitleFromParagraph(firstParagraph, out var extractedTitle)) {
            titleInlines = extractedTitle;
            blocks.RemoveAt(0);
        }

        callout = new CalloutBlock(kind, titleInlines, blocks);
        return true;
    }

    private static bool IsCalloutTitleExplicit(IElement element) {
        string? explicitAttribute = element.GetAttribute("data-omd-callout-title-explicit");
        if (string.IsNullOrWhiteSpace(explicitAttribute)) {
            return true;
        }

        if (bool.TryParse(explicitAttribute, out bool explicitTitle)) {
            return explicitTitle;
        }

        return true;
    }

    private static bool TryExtractCalloutTitleFromParagraph(ParagraphBlock paragraph, out InlineSequence titleInlines) {
        titleInlines = new InlineSequence();
        if (paragraph == null) {
            return false;
        }

        string markdown = paragraph.Inlines.RenderMarkdown().Trim();
        if (!IsStrongOnlyMarkdown(markdown)) {
            return false;
        }

        string inner = markdown.Substring(2, markdown.Length - 4).Trim();
        if (inner.Length == 0) {
            return false;
        }

        titleInlines = ParseInlines(inner);
        return true;
    }

    private static bool IsStrongOnlyParagraph(ParagraphBlock paragraph) {
        if (paragraph == null) {
            return false;
        }

        return IsStrongOnlyMarkdown(paragraph.Inlines.RenderMarkdown().Trim());
    }

    private static bool IsStrongOnlyMarkdown(string markdown) {
        if (markdown.Length < 4
            || !markdown.StartsWith("**", StringComparison.Ordinal)
            || !markdown.EndsWith("**", StringComparison.Ordinal)) {
            return false;
        }

        return true;
    }

    private static CodeBlock ConvertPreElement(IElement element) {
        var codeElement = element.QuerySelector("code");
        string language = string.Empty;
        if (codeElement != null) {
            language = ExtractCodeLanguage(codeElement.GetAttribute("class"));
        }

        string content = codeElement?.TextContent ?? element.TextContent ?? string.Empty;
        content = content.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
        return new CodeBlock(language, content);
    }

    private static string ExtractCodeLanguage(string? classValue) {
        if (string.IsNullOrWhiteSpace(classValue)) {
            return string.Empty;
        }

        foreach (string token in classValue!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (token.StartsWith("language-", StringComparison.OrdinalIgnoreCase)) {
                return token.Substring("language-".Length);
            }
            if (token.StartsWith("lang-", StringComparison.OrdinalIgnoreCase)) {
                return token.Substring("lang-".Length);
            }
        }

        return string.Empty;
    }

}
