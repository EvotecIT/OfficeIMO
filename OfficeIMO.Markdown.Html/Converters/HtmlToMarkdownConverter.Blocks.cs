using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static readonly HashSet<string> s_BlockTags = new(StringComparer.OrdinalIgnoreCase) {
        "ADDRESS", "ARTICLE", "ASIDE", "BLOCKQUOTE", "BODY", "DETAILS", "DIV", "DL", "FIGURE",
        "FOOTER", "FORM", "H1", "H2", "H3", "H4", "H5", "H6", "HEADER", "HR", "LI", "MAIN",
        "NAV", "OL", "P", "PICTURE", "PRE", "SECTION", "TABLE", "UL"
    };

    private static readonly HashSet<string> s_InlineTags = new(StringComparer.OrdinalIgnoreCase) {
        "A", "ABBR", "B", "BDI", "BDO", "BR", "BUTTON", "CITE", "CODE", "DATA", "DEL", "DFN",
        "EM", "I", "IMG", "INPUT", "INS", "KBD", "LABEL", "MARK", "Q", "RP", "RT", "RTC", "RUBY",
        "S", "SAMP", "SMALL", "SPAN", "STRONG", "SUB", "SUP", "TIME", "U", "VAR", "WBR"
    };

    private static List<IMarkdownBlock> ConvertNodesToBlocks(IEnumerable<INode> nodes, ConversionContext context) {
        var blocks = new List<IMarkdownBlock>();
        var inlineBuffer = new List<INode>();

        void FlushInlineParagraph() {
            if (inlineBuffer.Count == 0) {
                return;
            }

            InlineSequence inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(inlineBuffer, context));
            inlineBuffer.Clear();
            if (!HasVisibleInlineContent(inlineSequence)) {
                return;
            }

            blocks.Add(new ParagraphBlock(inlineSequence));
        }

        foreach (var node in nodes) {
            if (node is IComment) {
                continue;
            }

            if (node is IElement element && ShouldIgnoreElement(element, context)) {
                continue;
            }

            if (node is IElement blockElement && ShouldTreatAsBlockElement(blockElement, context)) {
                FlushInlineParagraph();
                blocks.AddRange(ConvertElementToBlocks(blockElement, context));
                continue;
            }

            inlineBuffer.Add(node);
        }

        FlushInlineParagraph();
        return blocks;
    }

    private static bool IsBlockElement(IElement element, ConversionContext context) {
        return s_BlockTags.Contains(GetEffectiveTagName(element, context));
    }

    private static bool IsInlineElement(IElement element, ConversionContext context) {
        return s_InlineTags.Contains(GetEffectiveTagName(element, context));
    }

    private static bool ShouldTreatAsBlockElement(IElement element, ConversionContext context) {
        if (context.Footnotes.TryGetDefinitionLabel(element, out _)
            || context.Footnotes.ShouldConvertContainer(element)
            || HtmlAccessibilitySemantics.TryGetHeadingLevel(element, out _)) {
            return true;
        }

        if (IsBlockElement(element, context)) {
            return true;
        }

        if (CanConvertAnchorToLinkedImageBlock(element, context)) {
            return true;
        }

        if (HasDirectBlockChildren(element, context)) {
            return true;
        }

        if (!IsInlineElement(element, context)) {
            if (IsVisualContractElement(element)
                || (context.Options.PreserveUnsupportedBlocks
                    && context.Options.UnknownBlockHandling == HtmlUnknownTagHandling.Preserve)) {
                return true;
            }

            if (HasVisibleInlineSibling(element, context)) {
                return false;
            }

            return context.Options.UnknownBlockHandling != HtmlUnknownTagHandling.Preserve;
        }

        return false;
    }

    private static bool IsVisualContractElement(IElement element) {
        if (element == null) {
            return false;
        }

        var attributes = new List<KeyValuePair<string, string?>>();
        foreach (var attribute in element.Attributes) {
            attributes.Add(new KeyValuePair<string, string?>(attribute.Name, attribute.Value));
        }

        return MarkdownVisualElementContract.TryParse(attributes, out _);
    }

    private static bool HasVisibleInlineSibling(IElement element, ConversionContext context) {
        if (element == null || element.ParentElement == null) {
            return false;
        }

        foreach (var sibling in element.ParentElement.ChildNodes) {
            if (ReferenceEquals(sibling, element)) {
                continue;
            }

            if (IsVisibleInlineFlowNode(sibling, context)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsVisibleInlineFlowNode(INode node, ConversionContext context) {
        switch (node) {
            case IText text:
                return !string.IsNullOrWhiteSpace(text.Data);
            case IElement element:
                if (ShouldIgnoreElement(element, context) || IsBlockElement(element, context)) {
                    return false;
                }

                return !HasDirectBlockChildren(element, context);
            default:
                return false;
        }
    }

    private static bool CanConvertAnchorToLinkedImageBlock(IElement element, ConversionContext context) {
        if (element == null
            || context == null
            || !HasEffectiveTagName(element, context, "A")) {
            return false;
        }

        string href = ResolveUrl(element.GetAttribute("href"), context);
        if (string.IsNullOrWhiteSpace(href)) {
            return false;
        }

        if (!TryResolveAnchorMediaElement(element, context, out var mediaElement)) {
            return false;
        }

        if (HasEffectiveTagName(mediaElement, context, "IMG")) {
            return CanCreateImageBlockWithoutSideEffects(mediaElement, context);
        }

        return HasEffectiveTagName(mediaElement, context, "PICTURE")
               && CanCreatePictureImageBlockWithoutSideEffects(mediaElement, context);
    }

    private static bool HasDirectBlockChildren(IElement element, ConversionContext context) {
        foreach (var child in element.Children) {
            if (ShouldTreatAsBlockElement(child, context)) {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<IMarkdownBlock> ConvertElementToBlocks(IElement element, ConversionContext context) {
        if (IsPassThroughTag(element, context)) {
            return new IMarkdownBlock[] { new HtmlRawBlock(NormalizeRawElement(element, context)) };
        }

        if (TryConvertVisualContractElement(element, context, out var visualBlock)) {
            return new IMarkdownBlock[] { visualBlock };
        }

        if (TryConvertMermaidElement(element, out var mermaidBlock)) {
            return new IMarkdownBlock[] { mermaidBlock };
        }

        if (TryConvertMathElement(element, out var mathBlock)) {
            return new IMarkdownBlock[] { mathBlock };
        }

        if (TryConvertConfiguredElementConverters(element, context, out var customBlocks)) {
            return customBlocks;
        }

        if (TryConvertFootnoteElement(element, context, out IReadOnlyList<IMarkdownBlock> footnoteBlocks)) {
            return footnoteBlocks;
        }

        if (HtmlAccessibilitySemantics.TryGetHeadingLevel(element, out int accessibleHeadingLevel)) {
            return new IMarkdownBlock[] { ConvertHeadingElement(element, accessibleHeadingLevel, context) };
        }

        string tag = GetEffectiveTagName(element, context);
        switch (tag) {
            case "P":
                return ConvertParagraphElement(element, context);
            case "H1":
            case "H2":
            case "H3":
            case "H4":
            case "H5":
            case "H6":
                return new IMarkdownBlock[] { ConvertHeadingElement(element, int.Parse(tag.Substring(1), System.Globalization.CultureInfo.InvariantCulture), context) };
            case "UL":
                return new IMarkdownBlock[] { ConvertListElement(element, ordered: false, context) };
            case "OL":
                return new IMarkdownBlock[] { ConvertListElement(element, ordered: true, context) };
            case "BLOCKQUOTE":
                return new IMarkdownBlock[] { ConvertBlockquoteElement(element, context) };
            case "PRE":
                return new IMarkdownBlock[] { ConvertPreElement(element) };
            case "TABLE":
                return new IMarkdownBlock[] { ConvertTableElement(element, context) };
            case "HR":
                return new IMarkdownBlock[] { new HorizontalRuleBlock() };
            case "IMG":
                return ConvertImageElement(element, context);
            case "PICTURE":
                return ConvertPictureElement(element, context);
            case "FIGURE":
                return ConvertFigureElement(element, context);
            case "AUDIO":
            case "VIDEO":
                return ConvertMediaElement(element, context);
            case "A":
                if (TryCreateLinkedImageBlockFromAnchor(element, context, out var linkedImage)) {
                    return new IMarkdownBlock[] { linkedImage };
                }

                if (HasRejectedHref(element, context)) {
                    var unwrappedBlocks = ConvertNodesToBlocks(element.ChildNodes, context).ToList();
                    return unwrappedBlocks;
                }

                if (TryConvertRejectedAnchorMedia(element, context, out var anchorMediaBlocks)) {
                    return anchorMediaBlocks;
                }

                if (context.Options.PreserveUnsupportedBlocks) {
                    return new IMarkdownBlock[] { new HtmlRawBlock(NormalizeRawElement(element, context)) };
                }

                var anchorInline = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
                if (!HasVisibleInlineContent(anchorInline)) {
                    return Array.Empty<IMarkdownBlock>();
                }

                return new IMarkdownBlock[] { new ParagraphBlock(anchorInline) };
            case "DETAILS":
                return new IMarkdownBlock[] { ConvertDetailsElement(element, context) };
            case "DL":
                return new IMarkdownBlock[] { ConvertDefinitionListElement(element, context) };
            case "DIV":
            case "SECTION":
            case "ARTICLE":
            case "MAIN":
            case "HEADER":
            case "FOOTER":
            case "NAV":
            case "ASIDE":
            case "FORM":
            case "ADDRESS":
            case "BODY":
                if (HasDirectBlockChildren(element, context)) {
                    return ConvertNodesToBlocks(element.ChildNodes, context);
                }

                var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
                if (!HasVisibleInlineContent(inlineSequence)) {
                    return Array.Empty<IMarkdownBlock>();
                }

                return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
            default:
                return ConvertUnknownElementToBlocks(element, context);
        }
    }

    private static bool HasRejectedHref(IElement element, ConversionContext context) {
        if (element == null || context == null || !HasEffectiveTagName(element, context, "A")) {
            return false;
        }

        string? href = element.GetAttribute("href");
        return !string.IsNullOrWhiteSpace(href) && string.IsNullOrWhiteSpace(ResolveUrl(href, context));
    }

    private static bool TryConvertRejectedAnchorMedia(IElement element, ConversionContext context, out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (element == null
            || context == null
            || !HasEffectiveTagName(element, context, "A")
            || !TryResolveAnchorMediaElement(element, context, out var mediaElement)) {
            return false;
        }

        if (!HasRejectedMediaSourceCandidate(mediaElement, context)
            && !HasBlockedBase64MediaSourceCandidate(mediaElement, context)) {
            return false;
        }

        blocks = ConvertNodesToBlocks(new INode[] { mediaElement }, context);
        return true;
    }

    private static bool TryConvertConfiguredElementConverters(IElement element, ConversionContext context, out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (element == null || context == null || context.Options.ElementBlockConverters.Count == 0) {
            return false;
        }

        var conversionContext = new HtmlElementBlockConversionContext(
            element,
            context.Options,
            nodes => ConvertNodesToBlocks(nodes, context),
            nodes => NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(nodes, context)),
            NormalizeBlockText);

        for (int i = 0; i < context.Options.ElementBlockConverters.Count; i++) {
            var converter = context.Options.ElementBlockConverters[i];
            if (converter == null) {
                continue;
            }

            var converted = converter.TryConvert(conversionContext);
            if (converted != null) {
                blocks = converted;
                return true;
            }
        }

        return false;
    }

    private static bool TryConvertVisualContractElement(IElement element, ConversionContext context, out SemanticFencedBlock visualBlock) {
        visualBlock = null!;
        if (element == null || context == null) {
            return false;
        }

        var attributes = new List<KeyValuePair<string, string?>>();
        foreach (var attribute in element.Attributes) {
            attributes.Add(new KeyValuePair<string, string?>(attribute.Name, attribute.Value));
        }

        if (!MarkdownVisualElementContract.TryParse(attributes, out var visualElement)) {
            return false;
        }

        var payload = visualElement!.TryDecodePayload();
        if (payload == null) {
            return false;
        }

        var roundTripContext = new MarkdownVisualElementRoundTripContext(
            element.TagName,
            visualElement,
            payload,
            TryReadVisualContractCaption(element));

        for (int i = 0; i < context.Options.VisualElementRoundTripHints.Count; i++) {
            var hint = context.Options.VisualElementRoundTripHints[i];
            if (hint == null) {
                continue;
            }

            var hintedBlock = hint.TryCreateBlock(roundTripContext);
            if (hintedBlock != null) {
                visualBlock = hintedBlock;
                return true;
            }
        }

        visualBlock = roundTripContext.CreateBlock();
        return true;
    }

    private static string? TryReadVisualContractCaption(IElement element) {
        if (element == null) {
            return null;
        }

        if (!element.TagName.Equals("FIGURE", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        foreach (var child in element.Children) {
            if (!child.TagName.Equals("FIGCAPTION", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var caption = NormalizeBlockText(child.TextContent);
            return caption.Length == 0 ? null : caption;
        }

        return null;
    }

    private static string BuildFenceInfoString(MarkdownVisualElement visualElement) {
        if (visualElement == null || string.IsNullOrWhiteSpace(visualElement.FenceLanguage)) {
            return string.Empty;
        }

        return MarkdownVisualElementRoundTripContext.BuildFenceInfoString(visualElement);
    }

    private static bool TryConvertMermaidElement(IElement element, out SemanticFencedBlock mermaidBlock) {
        mermaidBlock = null!;
        if (element == null
            || !element.TagName.Equals("PRE", StringComparison.OrdinalIgnoreCase)
            || !element.ClassList.Contains("mermaid")) {
            return false;
        }

        var content = (element.TextContent ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .TrimEnd('\n');
        if (string.IsNullOrWhiteSpace(content)) {
            return false;
        }

        mermaidBlock = new SemanticFencedBlock(MarkdownSemanticKinds.Mermaid, "mermaid", content);
        return true;
    }

    private static bool TryConvertMathElement(IElement element, out SemanticFencedBlock mathBlock) {
        mathBlock = null!;
        if (element == null
            || !element.TagName.Equals("DIV", StringComparison.OrdinalIgnoreCase)
            || !element.ClassList.Contains("omd-math")) {
            return false;
        }

        var content = (element.TextContent ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
        if (!TryExtractDisplayMathContent(content, out var mathContent)) {
            return false;
        }

        mathBlock = new SemanticFencedBlock(MarkdownSemanticKinds.Math, "math", mathContent);
        return true;
    }

    private static bool TryExtractDisplayMathContent(string content, out string mathContent) {
        mathContent = string.Empty;
        if (string.IsNullOrWhiteSpace(content)) {
            return false;
        }

        if (content.StartsWith("$$", StringComparison.Ordinal)
            && content.EndsWith("$$", StringComparison.Ordinal)
            && content.Length >= 4) {
            mathContent = content.Substring(2, content.Length - 4).Trim('\r', '\n');
            return mathContent.Length > 0;
        }

        return false;
    }

}
