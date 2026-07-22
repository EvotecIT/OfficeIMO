using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static string GetEffectiveTagName(IElement element, ConversionContext? context) {
        string tagName = element?.TagName ?? string.Empty;
        if (context?.Options.TagAliases == null || context.Options.TagAliases.Count == 0) {
            return tagName;
        }

        if (!context.Options.TagAliases.TryGetValue(tagName, out string? alias)
            && !context.Options.TagAliases.TryGetValue(tagName.ToLowerInvariant(), out alias)) {
            return tagName;
        }

        return string.IsNullOrWhiteSpace(alias)
            ? tagName
            : alias.Trim().ToUpperInvariant();
    }

    private static bool HasEffectiveTagName(IElement element, ConversionContext? context, string tagName) {
        return string.Equals(GetEffectiveTagName(element, context), tagName, StringComparison.OrdinalIgnoreCase);
    }

    private static IElement? FindFirstDescendantByEffectiveTagName(IElement element, ConversionContext? context, string tagName) {
        if (element == null) {
            return null;
        }

        foreach (var descendant in element.QuerySelectorAll("*")) {
            if (HasEffectiveTagName(descendant, context, tagName)) {
                return descendant;
            }
        }

        return null;
    }

    private static IElement? FindDirectChildByEffectiveTagName(IElement element, ConversionContext? context, string tagName) {
        if (element == null) {
            return null;
        }

        foreach (var child in element.Children) {
            if (HasEffectiveTagName(child, context, tagName)) {
                return child;
            }
        }

        return null;
    }

    private static bool IsPassThroughTag(IElement element, ConversionContext? context) {
        if (element == null || context?.Options.PassThroughTags == null || context.Options.PassThroughTags.Count == 0) {
            return false;
        }

        string originalTag = element.TagName;
        string effectiveTag = GetEffectiveTagName(element, context);
        return context.Options.PassThroughTags.Contains(originalTag)
               || context.Options.PassThroughTags.Contains(effectiveTag);
    }

    private static IEnumerable<IMarkdownBlock> ConvertUnknownElementToBlocks(IElement element, ConversionContext context) {
        switch (context.Options.UnknownBlockHandling) {
            case HtmlUnknownTagHandling.Drop:
                return Array.Empty<IMarkdownBlock>();
            case HtmlUnknownTagHandling.Bypass:
                return ConvertUnknownElementChildrenToBlocks(element, context);
            case HtmlUnknownTagHandling.Raise:
                throw new NotSupportedException($"Unsupported HTML block element '{element.TagName}' cannot be converted to Markdown.");
            case HtmlUnknownTagHandling.Preserve:
            default:
                if (context.Options.PreserveUnsupportedBlocks) {
                    return new IMarkdownBlock[] { new HtmlRawBlock(NormalizeRawElement(element, context)) };
                }

                return ConvertUnknownElementChildrenToBlocks(element, context);
        }
    }

    private static IEnumerable<IMarkdownBlock> ConvertUnknownElementChildrenToBlocks(IElement element, ConversionContext context) {
        if (HasDirectBlockChildren(element, context)) {
            return ConvertNodesToBlocks(element.ChildNodes, context);
        }

        var fallbackInline = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
        if (!HasVisibleInlineContent(fallbackInline)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(fallbackInline) };
    }

    private static void AppendUnknownInlineElement(InlineSequence sequence, IElement element, ConversionContext? context) {
        if (context == null) {
            AppendInlineElementChildren(sequence, element, context);
            return;
        }

        if (!IsInlineElement(element, context) && context.Options.UnknownInlineHandling == HtmlUnknownTagHandling.Preserve) {
            switch (context.Options.UnknownBlockHandling) {
                case HtmlUnknownTagHandling.Drop:
                    return;
                case HtmlUnknownTagHandling.Bypass:
                    AppendInlineElementChildren(sequence, element, context);
                    return;
                case HtmlUnknownTagHandling.Raise:
                    throw new NotSupportedException($"Unsupported HTML inline element '{element.TagName}' cannot be converted to Markdown.");
            }
        }

        switch (context.Options.UnknownInlineHandling) {
            case HtmlUnknownTagHandling.Drop:
                return;
            case HtmlUnknownTagHandling.Bypass:
                AppendInlineElementChildren(sequence, element, context);
                return;
            case HtmlUnknownTagHandling.Raise:
                throw new NotSupportedException($"Unsupported HTML inline element '{element.TagName}' cannot be converted to Markdown.");
            case HtmlUnknownTagHandling.Preserve:
            default:
                if (context.Options.PreserveUnsupportedInlineHtml) {
                    sequence.AddRaw(new HtmlRawInline(NormalizeRawElement(element, context)));
                    return;
                }

                AppendInlineElementChildren(sequence, element, context);
                return;
        }
    }

    private static string ConvertUnknownInlineElementToMarkdown(IElement element, ConversionContext? context) {
        if (context == null) {
            return ConvertInlineNodesToMarkdown(element.ChildNodes, context);
        }

        switch (context.Options.UnknownInlineHandling) {
            case HtmlUnknownTagHandling.Drop:
                return string.Empty;
            case HtmlUnknownTagHandling.Bypass:
                return ConvertInlineNodesToMarkdown(element.ChildNodes, context);
            case HtmlUnknownTagHandling.Raise:
                throw new NotSupportedException($"Unsupported HTML inline element '{element.TagName}' cannot be converted to Markdown.");
            case HtmlUnknownTagHandling.Preserve:
            default:
                return context.Options.PreserveUnsupportedInlineHtml
                    ? element.OuterHtml
                    : ConvertInlineNodesToMarkdown(element.ChildNodes, context);
        }
    }

    private static void AppendInlineElementChildren(InlineSequence sequence, IElement element, ConversionContext? context) {
        var childNodes = element.ChildNodes as IList<INode> ?? element.ChildNodes.ToList();
        for (int i = 0; i < childNodes.Count; i++) {
            bool trimEnd = NextVisibleInlineNodeIsBoundary(childNodes, i + 1, context);
            AppendInlineNode(sequence, childNodes[i], context, trimEnd);
        }
    }
}
