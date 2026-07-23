using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
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
            list.MarkerStyle = ParseOrderedListMarkerStyle(element.GetAttribute("type"));
            IElement[] itemElements = element.Children
                .Where(child => HasEffectiveTagName(child, context, "LI"))
                .ToArray();
            bool reversed = element.HasAttribute("reversed");
            list.Reversed = reversed;
            int currentValue = reversed ? itemElements.Length : 1;
            if (int.TryParse(
                    element.GetAttribute("start"),
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out int start)) {
                currentValue = start;
            }
            list.Start = currentValue;
            int step = reversed ? -1 : 1;

            foreach (IElement itemElement in itemElements) {
                if (int.TryParse(
                        itemElement.GetAttribute("value"),
                        System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out int itemValue)) {
                    currentValue = itemValue;
                }
                ListItem item = ConvertListItem(itemElement, context);
                item.MarkerText = currentValue.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".";
                list.Items.Add(item);
                currentValue += step;
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
                    item.NestedBlocks.Add(paragraph);
                }
            } else {
                encounteredNonParagraphBlock = true;
                item.NestedBlocks.Add(blocks[index]);
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
            quote.ChildBlocks.Add(block);
        }

        if (quote.ChildBlocks.Count == 0) {
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
        string language = ExtractCodeLanguage(codeElement?.GetAttribute("class"));
        if (language.Length == 0) language = ReadCodeLanguageAttribute(codeElement);
        if (language.Length == 0) language = ExtractCodeLanguage(element.GetAttribute("class"));
        if (language.Length == 0) language = ReadCodeLanguageAttribute(element);
        language = NormalizeCodeLanguage(language);

        string content = codeElement?.TextContent ?? element.TextContent ?? string.Empty;
        content = content.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
        return new CodeBlock(language, content);
    }

    private static MarkdownOrderedListMarkerStyle ParseOrderedListMarkerStyle(string? value) {
        switch (value?.Trim()) {
            case "a": return MarkdownOrderedListMarkerStyle.LowerAlpha;
            case "A": return MarkdownOrderedListMarkerStyle.UpperAlpha;
            case "i": return MarkdownOrderedListMarkerStyle.LowerRoman;
            case "I": return MarkdownOrderedListMarkerStyle.UpperRoman;
            default: return MarkdownOrderedListMarkerStyle.Decimal;
        }
    }

    private static string ReadCodeLanguageAttribute(IElement? element) {
        if (element == null) return string.Empty;
        foreach (string name in new[] { "data-language", "data-lang", "lang" }) {
            string? value = element.GetAttribute(name);
            if (!string.IsNullOrWhiteSpace(value)) return value!.Trim();
        }
        return string.Empty;
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

    private static string NormalizeCodeLanguage(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string candidate = value!.Trim();
        if (candidate.Length > 64) return string.Empty;
        for (int index = 0; index < candidate.Length; index++) {
            char character = candidate[index];
            if (!(char.IsLetterOrDigit(character)
                || character == '-'
                || character == '_'
                || character == '+'
                || character == '.'
                || character == '#')) {
                return string.Empty;
            }
        }
        return candidate;
    }

}
