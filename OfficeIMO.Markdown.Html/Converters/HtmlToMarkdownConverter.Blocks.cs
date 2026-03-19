using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
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

    private static bool IsBlockElement(IElement element) {
        return s_BlockTags.Contains(element.TagName);
    }

    private static bool IsInlineElement(IElement element) {
        return s_InlineTags.Contains(element.TagName);
    }

    private static bool ShouldTreatAsBlockElement(IElement element, ConversionContext context) {
        if (IsBlockElement(element)) {
            return true;
        }

        if (CanConvertAnchorToLinkedImageBlock(element, context)) {
            return true;
        }

        if (HasDirectBlockChildren(element, context)) {
            return true;
        }

        if (context.Options.PreserveUnsupportedBlocks && !IsInlineElement(element)) {
            return true;
        }

        return false;
    }

    private static bool CanConvertAnchorToLinkedImageBlock(IElement element, ConversionContext context) {
        if (element == null
            || context == null
            || !element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return TryCreateLinkedImageBlockFromAnchor(element, context, out _);
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

        string tag = element.TagName;
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
            case "A":
                if (TryCreateLinkedImageBlockFromAnchor(element, context, out var linkedImage)) {
                    return new IMarkdownBlock[] { linkedImage };
                }

                if (context.Options.PreserveUnsupportedBlocks) {
                    return new IMarkdownBlock[] { new HtmlRawBlock(element.OuterHtml) };
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
                if (context.Options.PreserveUnsupportedBlocks) {
                    return new IMarkdownBlock[] { new HtmlRawBlock(element.OuterHtml) };
                }

                if (HasDirectBlockChildren(element, context)) {
                    return ConvertNodesToBlocks(element.ChildNodes, context);
                }

                var fallbackInline = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
                if (!HasVisibleInlineContent(fallbackInline)) {
                    return Array.Empty<IMarkdownBlock>();
                }

                return new IMarkdownBlock[] { new ParagraphBlock(fallbackInline) };
        }
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

        bool isStandaloneMedia = onlyElement.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                                 || onlyElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
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

            foreach (var item in element.Children.Where(child => child.TagName.Equals("LI", StringComparison.OrdinalIgnoreCase))) {
                list.Items.Add(ConvertListItem(item, context));
            }

            return list;
        }

        var unordered = new UnorderedListBlock();
        foreach (var item in element.Children.Where(child => child.TagName.Equals("LI", StringComparison.OrdinalIgnoreCase))) {
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
                && childElement.TagName.Equals("INPUT", StringComparison.OrdinalIgnoreCase)
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

    private static TableBlock ConvertTableElement(IElement element, ConversionContext context) {
        var table = new TableBlock();
        bool headerWritten = false;
        var headerCells = new List<TableCell>();
        var rowCells = new List<IReadOnlyList<TableCell>>();

        foreach (var row in EnumerateTableRows(element)) {
            var cells = row.Children
                .Where(child => child.TagName.Equals("TH", StringComparison.OrdinalIgnoreCase) || child.TagName.Equals("TD", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (cells.Count == 0) {
                continue;
            }

            bool isHeaderRow = !headerWritten && cells.All(cell => cell.TagName.Equals("TH", StringComparison.OrdinalIgnoreCase));
            var renderedCells = new List<string>(cells.Count);
            var structuredCells = new List<TableCell>(cells.Count);
            foreach (var cell in cells) {
                var cellBlocks = ConvertTableCellToBlocks(cell, context);
                structuredCells.Add(new TableCell(cellBlocks));
                renderedCells.Add(RenderTableCellBlocksToMarkdown(cellBlocks));
                if (isHeaderRow) {
                    table.Alignments.Add(ParseAlignment(cell));
                }
            }

            if (isHeaderRow) {
                foreach (var value in renderedCells) {
                    table.Headers.Add(value);
                }
                headerCells.AddRange(structuredCells);
                headerWritten = true;
            } else {
                table.Rows.Add(renderedCells);
                rowCells.Add(structuredCells);
            }
        }

        if (!headerWritten && table.Rows.Count > 0) {
            var firstRow = table.Rows[0];
            table.Rows.RemoveAt(0);
            var firstStructuredRow = rowCells[0];
            rowCells.RemoveAt(0);
            foreach (var value in firstRow) {
                table.Headers.Add(value);
                table.Alignments.Add(ColumnAlignment.None);
            }
            headerCells.AddRange(firstStructuredRow);
        }

        table.SetStructuredCells(headerCells, rowCells, table.ComputeContentSignature());

        return table;
    }

    private static IEnumerable<IElement> EnumerateTableRows(IElement table) {
        foreach (var child in table.Children) {
            if (child.TagName.Equals("TR", StringComparison.OrdinalIgnoreCase)) {
                yield return child;
                continue;
            }

            if (!child.TagName.Equals("THEAD", StringComparison.OrdinalIgnoreCase)
                && !child.TagName.Equals("TBODY", StringComparison.OrdinalIgnoreCase)
                && !child.TagName.Equals("TFOOT", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            foreach (var row in child.Children.Where(static row => row.TagName.Equals("TR", StringComparison.OrdinalIgnoreCase))) {
                yield return row;
            }
        }
    }

    private static ColumnAlignment ParseAlignment(IElement cell) {
        var alignment = ParseAlignmentValue(cell.GetAttribute("align"));
        if (alignment != ColumnAlignment.None) {
            return alignment;
        }

        return ParseAlignmentValue(TryGetStyleDeclarationValue(cell.GetAttribute("style"), "text-align"));
    }

    private static ColumnAlignment ParseAlignmentValue(string? rawAlignment) {
        if (string.IsNullOrWhiteSpace(rawAlignment)) {
            return ColumnAlignment.None;
        }

        switch (rawAlignment!.Trim().ToLowerInvariant()) {
            case "left":
                return ColumnAlignment.Left;
            case "center":
                return ColumnAlignment.Center;
            case "right":
                return ColumnAlignment.Right;
            default:
                return ColumnAlignment.None;
        }
    }

    private static IReadOnlyList<IMarkdownBlock> ConvertTableCellToBlocks(IElement cell, ConversionContext context) {
        if (HasDirectBlockChildren(cell, context)) {
            return ConvertNodesToBlocks(cell.ChildNodes, context);
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(cell.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static string RenderTableCellBlocksToMarkdown(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        return new TableCell(blocks).Markdown.Replace("  \n", "<br>");
    }

    private static IEnumerable<IMarkdownBlock> ConvertImageElement(IElement element, ConversionContext context) {
        if (!TryCreateImageBlock(element, context, out var image)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { image };
    }

    private static IEnumerable<IMarkdownBlock> ConvertPictureElement(IElement element, ConversionContext context) {
        if (element == null) {
            return Array.Empty<IMarkdownBlock>();
        }

        string preferredSrc = ResolvePictureSource(element, context);
        var imageElement = element.QuerySelector("img");
        if (imageElement != null && TryCreateImageBlock(imageElement, context, out var imageBlock)) {
            imageBlock = !string.IsNullOrWhiteSpace(preferredSrc)
                ? CreateImageBlock(preferredSrc, imageElement, element, context)
                : CreateImageBlock(imageBlock.Path, imageElement, element, context);

            return new IMarkdownBlock[] { imageBlock };
        }

        if (string.IsNullOrWhiteSpace(preferredSrc)) {
            return context.Options.PreserveUnsupportedBlocks
                ? new IMarkdownBlock[] { new HtmlRawBlock(element.OuterHtml) }
                : Array.Empty<IMarkdownBlock>();
        }

        var pictureImage = new ImageBlock(preferredSrc, alt: null, title: null);
        ApplyPictureMetadata(element, pictureImage, null, context);
        return new IMarkdownBlock[] { pictureImage };
    }

    private static IEnumerable<IMarkdownBlock> ConvertFigureElement(IElement element, ConversionContext context) {
        var directCaption = element.Children.FirstOrDefault(child => child.TagName.Equals("FIGCAPTION", StringComparison.OrdinalIgnoreCase));
        var directMediaContainer = element.Children.FirstOrDefault(child => TryResolveFigureMediaElement(child, out _));

        if (directMediaContainer != null && TryResolveFigureMediaElement(directMediaContainer, out var directMedia)) {
            var figureBlocks = new List<IMarkdownBlock>();
            foreach (var child in element.ChildNodes) {
                if (ReferenceEquals(child, directCaption)) {
                    continue;
                }

                if (ReferenceEquals(child, directMediaContainer)) {
                    var mediaBlocks = ConvertFigureMediaElement(directMedia, context);
                    ApplyFigureCaptionToMedia(mediaBlocks, directCaption);
                    figureBlocks.AddRange(mediaBlocks);
                    continue;
                }

                figureBlocks.AddRange(ConvertNodesToBlocks(new[] { child }, context));
            }

            if (figureBlocks.Count > 0) {
                return figureBlocks;
            }
        }

        var imageElement = element.QuerySelector("img");
        if (imageElement == null) {
            var pictureElement = element.QuerySelector("picture");
            if (pictureElement != null) {
                var pictureBlocks = ConvertPictureElement(pictureElement, context).ToList();
                if (pictureBlocks.Count > 0) {
                    ApplyFigureCaptionToMedia(pictureBlocks, directCaption ?? element.QuerySelector("figcaption"));

                    return pictureBlocks;
                }
            }

            if (HasDirectBlockChildren(element, context)) {
                return ConvertNodesToBlocks(element.ChildNodes, context);
            }

            var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
            if (!HasVisibleInlineContent(inlineSequence)) {
                return Array.Empty<IMarkdownBlock>();
            }

            return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
        }

        var blocks = ConvertImageElement(imageElement, context).ToList();
        ApplyFigureCaptionToMedia(blocks, directCaption ?? element.QuerySelector("figcaption"));

        return blocks;
    }

    private static void ApplyFigureCaptionToMedia(IReadOnlyList<IMarkdownBlock> blocks, IElement? captionElement) {
        if (captionElement == null || blocks == null || blocks.Count != 1 || blocks[0] is not ImageBlock imageBlock) {
            return;
        }

        imageBlock.Caption = NormalizeBlockText(captionElement.TextContent);
    }

    private static bool IsLinkedFigureMediaAnchor(IElement element) {
        return TryResolveAnchorMediaElement(element, out _);
    }

    private static bool TryResolveFigureMediaElement(IElement element, out IElement mediaElement) {
        return TryResolvePureWrapperElement(
            element,
            candidate =>
                candidate.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                || candidate.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
                || IsLinkedFigureMediaAnchor(candidate),
            candidate => !candidate.TagName.Equals("FIGCAPTION", StringComparison.OrdinalIgnoreCase),
            out mediaElement);
    }

    private static bool TryResolveAnchorMediaElement(IElement element, out IElement mediaElement) {
        mediaElement = null!;
        if (element == null || !element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase) || HasVisibleOwnTextNodes(element)) {
            return false;
        }

        foreach (var childNode in element.ChildNodes) {
            switch (childNode) {
                case IComment:
                    continue;
                case IText textNode when string.IsNullOrWhiteSpace(textNode.Data):
                    continue;
                case IElement childElement when IsIgnorableMediaWrapperChild(childElement):
                    continue;
                case IElement childElement when childElement.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                    || childElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase):
                    mediaElement = childElement;
                    return true;
                case IElement childElement when !childElement.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)
                    && TryResolvePureWrapperElement(
                        childElement,
                        candidate =>
                            candidate.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                            || candidate.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase),
                        candidate => !candidate.TagName.Equals("A", StringComparison.OrdinalIgnoreCase),
                        out mediaElement):
                    return true;
                default:
                    return false;
            }
        }

        return false;
    }

    private static bool IsIgnorableMediaWrapperChild(IElement element) {
        if (element == null) {
            return false;
        }

        return element.TagName.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("SCRIPT", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("STYLE", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("TEMPLATE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolvePureWrapperElement(
        IElement element,
        Func<IElement, bool> terminalPredicate,
        Func<IElement, bool> canRecursePredicate,
        out IElement resolvedElement) {
        resolvedElement = null!;
        if (element == null) {
            return false;
        }

        if (terminalPredicate(element)) {
            resolvedElement = element;
            return true;
        }

        if (!canRecursePredicate(element) || HasVisibleOwnTextNodes(element)) {
            return false;
        }

        IElement? onlyChildElement = null;
        foreach (var childNode in element.ChildNodes) {
            switch (childNode) {
                case IComment:
                    continue;
                case IText textNode when string.IsNullOrWhiteSpace(textNode.Data):
                    continue;
                case IElement childElement when IsIgnorableMediaWrapperChild(childElement):
                    continue;
                case IElement childElement:
                    if (onlyChildElement != null) {
                        return false;
                    }

                    onlyChildElement = childElement;
                    break;
                default:
                    return false;
            }
        }

        return onlyChildElement != null
            && TryResolvePureWrapperElement(onlyChildElement, terminalPredicate, canRecursePredicate, out resolvedElement);
    }

    private static bool HasVisibleOwnTextNodes(IElement element) {
        foreach (var childNode in element.ChildNodes) {
            if (childNode is IText textNode && !string.IsNullOrWhiteSpace(textNode.Data)) {
                return true;
            }
        }

        return false;
    }

    private static List<IMarkdownBlock> ConvertFigureMediaElement(IElement element, ConversionContext context) {
        if (element.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            return ConvertPictureElement(element, context).ToList();
        }

        if (element.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)) {
            return ConvertImageElement(element, context).ToList();
        }

        if (TryCreateLinkedImageBlockFromAnchor(element, context, out var linkedImage)) {
            return new List<IMarkdownBlock> { linkedImage };
        }

        return ConvertNodesToBlocks(new[] { element }, context);
    }

    private static DetailsBlock ConvertDetailsElement(IElement element, ConversionContext context) {
        SummaryBlock? summary = null;
        var summaryElement = element.Children.FirstOrDefault(child => child.TagName.Equals("SUMMARY", StringComparison.OrdinalIgnoreCase));
        if (summaryElement != null) {
            summary = new SummaryBlock(NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(summaryElement.ChildNodes, context)));
        }

        var details = new DetailsBlock(summary, open: element.HasAttribute("open"));
        foreach (var child in element.ChildNodes) {
            if (ReferenceEquals(child, summaryElement)) {
                continue;
            }

            foreach (var block in ConvertNodesToBlocks(new[] { child }, context)) {
                details.Children.Add(block);
            }
        }

        return details;
    }

    private static DefinitionListBlock ConvertDefinitionListElement(IElement element, ConversionContext context) {
        var list = new DefinitionListBlock();
        var pendingTerms = new List<InlineSequence>();
        bool hasDefinitionsForCurrentGroup = false;

        foreach (var child in element.Children) {
            if (child.TagName.Equals("DT", StringComparison.OrdinalIgnoreCase)) {
                if (hasDefinitionsForCurrentGroup) {
                    pendingTerms.Clear();
                    hasDefinitionsForCurrentGroup = false;
                }

                var term = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(child.ChildNodes, context));
                if (HasVisibleInlineContent(term)) {
                    pendingTerms.Add(term);
                }
                continue;
            }

            if (child.TagName.Equals("DD", StringComparison.OrdinalIgnoreCase) && pendingTerms.Count > 0) {
                foreach (var term in pendingTerms) {
                    list.AddEntry(new DefinitionListEntry(
                        term,
                        ConvertDefinitionValueToBlocks(child, context)));
                }
                hasDefinitionsForCurrentGroup = true;
            }
        }

        return list;
    }

    private static IReadOnlyList<IMarkdownBlock> ConvertDefinitionValueToBlocks(IElement element, ConversionContext context) {
        if (HasDirectBlockChildren(element, context)) {
            return ConvertNodesToBlocks(element.ChildNodes, context);
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static InlineSequence NormalizeInlineSequenceForBlock(InlineSequence? source) {
        return source ?? new InlineSequence { AutoSpacing = false };
    }

    private static bool HasVisibleInlineContent(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        foreach (var node in sequence.Nodes) {
            switch (node) {
                case null:
                    continue;
                case TextRun textRun when string.IsNullOrWhiteSpace(textRun.Text):
                    continue;
                default:
                    return true;
            }
        }

        return false;
    }

    private static string ResolveImageSource(IElement element, ConversionContext context) {
        return ResolveImageSource(element, context, allowParentPictureFallback: true);
    }

    private static string ResolveDirectImageSource(IElement element, ConversionContext context) {
        return ResolveImageSource(element, context, allowParentPictureFallback: false);
    }

    private static string ResolveImageSource(IElement element, ConversionContext context, bool allowParentPictureFallback) {
        string[] lazySourceAttributes = new[] { "data-src", "data-original", "data-original-src", "data-lazy-src" };
        for (int i = 0; i < lazySourceAttributes.Length; i++) {
            string resolved = ResolveUrl(element.GetAttribute(lazySourceAttributes[i]), context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        string[] sourceAttributes = new[] { "src" };
        for (int i = 0; i < sourceAttributes.Length; i++) {
            string resolved = ResolveUrl(element.GetAttribute(sourceAttributes[i]), context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        string srcSetResolved = ResolveUrlFromSrcSetAttributes(element, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
        if (!string.IsNullOrWhiteSpace(srcSetResolved)) {
            return srcSetResolved;
        }

        return allowParentPictureFallback
            && element.ParentElement != null
            && element.ParentElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
            ? ResolvePictureSource(element.ParentElement, context)
            : string.Empty;
    }

    private static string ResolvePictureSource(IElement pictureElement, ConversionContext context) {
        if (pictureElement == null) {
            return string.Empty;
        }

        foreach (var child in pictureElement.Children) {
            if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string resolved = ResolveUrlFromSrcSetAttributes(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }

            resolved = ResolveUrlAttributes(child, context, "src", "data-src", "data-original-src", "data-lazy-src");
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    private static string ResolveUrlFromSrcSet(string? rawSrcSet, ConversionContext context) {
        return GetFirstResolvedSrcSetCandidate(rawSrcSet, context).url;
    }

    private static string ResolveNormalizedSrcSet(string? rawSrcSet, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(rawSrcSet)) {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(rawSrcSet)) {
            string resolved = ResolveUrl(candidate.Url, context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                parts.Add(string.IsNullOrWhiteSpace(candidate.Descriptor) ? resolved : resolved + " " + candidate.Descriptor);
            }
        }

        return string.Join(", ", parts);
    }

    private static (string url, string descriptor) GetFirstResolvedSrcSetCandidate(string? rawSrcSet, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(rawSrcSet)) {
            return (string.Empty, string.Empty);
        }

        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(rawSrcSet)) {
            string resolved = ResolveUrl(candidate.Url, context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return (resolved, candidate.Descriptor);
            }
        }

        return (string.Empty, string.Empty);
    }

    private static string ResolveUrlFromSrcSetAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = ResolveUrlFromSrcSet(element.GetAttribute(attributeNames[i]), context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    private static string ResolveNormalizedSrcSetAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = ResolveNormalizedSrcSet(element.GetAttribute(attributeNames[i]), context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    private static string ResolveUrlAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = ResolveUrl(element.GetAttribute(attributeNames[i]), context);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    private static void ApplyImageDimensions(IElement element, ImageBlock image) {
        if (TryParseImageDimension(element.GetAttribute("width"), out double width)
            || TryParseStyleDimension(element.GetAttribute("style"), "width", out width)) {
            image.Width = width;
        }

        if (TryParseImageDimension(element.GetAttribute("height"), out double height)
            || TryParseStyleDimension(element.GetAttribute("style"), "height", out height)) {
            image.Height = height;
        }
    }

    private static bool TryCreateImageBlock(IElement element, ConversionContext context, out ImageBlock image) {
        image = null!;
        string src = ResolveImageSource(element, context);
        if ((string.IsNullOrWhiteSpace(src) || IsLikelyPlaceholderImageSource(src))
            && TryCreateImageBlockFromNoscriptFallback(element, context, out image)) {
            return true;
        }

        if (string.IsNullOrWhiteSpace(src)) {
            return false;
        }

        image = CreateImageBlock(src, element);
        return true;
    }

    private static ImageBlock CreateImageBlock(string src, IElement metadataElement, IElement? pictureElement = null, ConversionContext? context = null) {
        var image = new ImageBlock(
            src,
            metadataElement.GetAttribute("alt"),
            metadataElement.GetAttribute("title"));
        ApplyImageDimensions(metadataElement, image);
        if (pictureElement != null && context != null) {
            ApplyPictureMetadata(pictureElement, image, metadataElement, context);
        }
        return image;
    }

    private static bool TryCreateImageBlockFromNoscriptFallback(IElement element, ConversionContext context, out ImageBlock image) {
        image = null!;
        if (!TryResolveAssociatedNoscriptMediaElement(element, out var fallbackMediaElement)) {
            return false;
        }

        if (fallbackMediaElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            var pictureImage = ConvertPictureElement(fallbackMediaElement, context).OfType<ImageBlock>().FirstOrDefault();
            if (pictureImage == null) {
                return false;
            }

            image = MergeImageMetadata(element, pictureImage, fallbackMediaElement.QuerySelector("img"));
            return true;
        }

        string fallbackSrc = ResolveImageSource(fallbackMediaElement, context);
        if (string.IsNullOrWhiteSpace(fallbackSrc)) {
            return false;
        }

        image = CreateMergedImageBlock(fallbackSrc, element, fallbackMediaElement);
        return true;
    }

    private static ImageBlock MergeImageMetadata(IElement preferredElement, ImageBlock fallbackImage, IElement? fallbackMetadataElement) {
        var merged = new ImageBlock(
            fallbackImage.Path,
            !string.IsNullOrWhiteSpace(preferredElement.GetAttribute("alt")) ? preferredElement.GetAttribute("alt") : fallbackImage.Alt,
            !string.IsNullOrWhiteSpace(preferredElement.GetAttribute("title")) ? preferredElement.GetAttribute("title") : fallbackImage.Title,
            fallbackImage.Width,
            fallbackImage.Height,
            fallbackImage.LinkUrl,
            fallbackImage.LinkTitle,
            fallbackImage.LinkTarget,
            fallbackImage.LinkRel) {
            Caption = fallbackImage.Caption,
            PictureFallbackPath = fallbackImage.PictureFallbackPath
        };
        CopyPictureSources(fallbackImage.PictureSources, merged.PictureSources);

        ApplyImageDimensions(preferredElement, merged);
        if ((merged.Width == null || merged.Height == null) && fallbackMetadataElement != null) {
            ApplyMissingImageDimensions(fallbackMetadataElement, merged);
        }

        return merged;
    }

    private static ImageBlock CreateMergedImageBlock(string src, IElement preferredMetadataElement, IElement fallbackMetadataElement) {
        var image = new ImageBlock(
            src,
            !string.IsNullOrWhiteSpace(preferredMetadataElement.GetAttribute("alt")) ? preferredMetadataElement.GetAttribute("alt") : fallbackMetadataElement.GetAttribute("alt"),
            !string.IsNullOrWhiteSpace(preferredMetadataElement.GetAttribute("title")) ? preferredMetadataElement.GetAttribute("title") : fallbackMetadataElement.GetAttribute("title"));
        ApplyImageDimensions(preferredMetadataElement, image);
        ApplyMissingImageDimensions(fallbackMetadataElement, image);
        return image;
    }

    private static void ApplyPictureMetadata(IElement pictureElement, ImageBlock image, IElement? fallbackImageElement, ConversionContext context) {
        if (pictureElement == null || image == null || context == null) {
            return;
        }

        image.PictureSources.Clear();
        foreach (var source in CollectPictureSources(pictureElement, context)) {
            image.PictureSources.Add(source);
        }

        string fallbackPath = fallbackImageElement == null
            ? string.Empty
            : ResolveDirectImageSource(fallbackImageElement, context);
        image.PictureFallbackPath = string.IsNullOrWhiteSpace(fallbackPath) ? null : fallbackPath;
    }

    private static List<ImagePictureSource> CollectPictureSources(IElement pictureElement, ConversionContext context) {
        var sources = new List<ImagePictureSource>();
        foreach (var child in pictureElement.Children) {
            if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string resolvedSrcSet = ResolveNormalizedSrcSetAttributes(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            string resolved = ResolveUrlFromSrcSetAttributes(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            if (string.IsNullOrWhiteSpace(resolved)) {
                resolved = ResolveUrlAttributes(child, context, "src", "data-src", "data-original-src", "data-lazy-src");
            }

            if (string.IsNullOrWhiteSpace(resolved)) {
                continue;
            }

            sources.Add(new ImagePictureSource(
                resolved,
                child.GetAttribute("media"),
                child.GetAttribute("type"),
                child.GetAttribute("sizes"),
                resolvedSrcSet));
        }

        return sources;
    }

    private static void CopyPictureSources(IEnumerable<ImagePictureSource> sourceItems, IList<ImagePictureSource> targetItems) {
        if (sourceItems == null || targetItems == null) {
            return;
        }

        targetItems.Clear();
        foreach (var source in sourceItems) {
            if (source == null || string.IsNullOrWhiteSpace(source.Path)) {
                continue;
            }

            targetItems.Add(new ImagePictureSource(source.Path, source.Media, source.Type, source.Sizes, source.SrcSet));
        }
    }

    private static void ApplyMissingImageDimensions(IElement element, ImageBlock image) {
        if (image.Width == null
            && (TryParseImageDimension(element.GetAttribute("width"), out double width)
                || TryParseStyleDimension(element.GetAttribute("style"), "width", out width))) {
            image.Width = width;
        }

        if (image.Height == null
            && (TryParseImageDimension(element.GetAttribute("height"), out double height)
                || TryParseStyleDimension(element.GetAttribute("style"), "height", out height))) {
            image.Height = height;
        }
    }

    private static bool TryResolveAssociatedNoscriptMediaElement(IElement element, out IElement mediaElement) {
        mediaElement = null!;
        foreach (var noscriptElement in EnumerateAssociatedNoscriptElements(element)) {
            if (TryResolveNoscriptMediaElement(noscriptElement, out mediaElement)) {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<IElement> EnumerateAssociatedNoscriptElements(IElement element) {
        var visited = new HashSet<IElement>();
        IElement? current = element.ParentElement;
        int depth = 0;
        while (current != null && depth < 3) {
            foreach (var child in current.Children) {
                if (child.TagName.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase) && visited.Add(child)) {
                    yield return child;
                }
            }

            if (!IsPotentialMediaFallbackContainer(current)) {
                yield break;
            }

            current = current.ParentElement;
            depth++;
        }
    }

    private static bool IsPotentialMediaFallbackContainer(IElement element) {
        if (element == null) {
            return false;
        }

        return element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("DIV", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("SPAN", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("FIGURE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolveNoscriptMediaElement(IElement noscriptElement, out IElement mediaElement) {
        mediaElement = null!;
        foreach (string html in EnumerateNoscriptHtmlCandidates(noscriptElement)) {
            var parser = new AngleSharp.Html.Parser.HtmlParser();
            var document = parser.ParseDocument($"<body>{html}</body>");
            IElement? parsedMediaElement = document.QuerySelector("picture") ?? document.QuerySelector("img");
            if (parsedMediaElement != null) {
                mediaElement = parsedMediaElement;
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<string> EnumerateNoscriptHtmlCandidates(IElement noscriptElement) {
        if (noscriptElement == null) {
            yield break;
        }

        string innerHtml = noscriptElement.InnerHtml;
        if (!string.IsNullOrWhiteSpace(innerHtml)) {
            yield return innerHtml;
        }

        string textContent = noscriptElement.TextContent;
        if (!string.IsNullOrWhiteSpace(textContent) && !string.Equals(textContent, innerHtml, StringComparison.Ordinal)) {
            yield return textContent;
        }
    }

    private static bool IsLikelyPlaceholderImageSource(string? source) {
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        string value = source!.Trim();
        return value.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase)
               || value.Contains("transparent.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("spacer.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("blank.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("pixel.gif", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCreateLinkedImageBlockFromAnchor(IElement anchorElement, ConversionContext context, out ImageBlock image) {
        image = null!;
        if (anchorElement == null || !anchorElement.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        string href = ResolveUrl(anchorElement.GetAttribute("href"), context);
        if (string.IsNullOrWhiteSpace(href)) {
            return false;
        }

        if (!TryResolveAnchorMediaElement(anchorElement, out var mediaElement)) {
            return false;
        }

        if (mediaElement.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
            && TryCreateImageBlock(mediaElement, context, out image)) {
            image.LinkUrl = href;
            image.LinkTitle = anchorElement.GetAttribute("title");
            image.LinkTarget = anchorElement.GetAttribute("target");
            image.LinkRel = anchorElement.GetAttribute("rel");
            return true;
        }

        if (!mediaElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        var pictureImage = ConvertPictureElement(mediaElement, context).OfType<ImageBlock>().FirstOrDefault();
        if (pictureImage == null) {
            return false;
        }

        pictureImage.LinkUrl = href;
        pictureImage.LinkTitle = anchorElement.GetAttribute("title");
        pictureImage.LinkTarget = anchorElement.GetAttribute("target");
        pictureImage.LinkRel = anchorElement.GetAttribute("rel");
        image = pictureImage;
        return true;
    }

    private static bool TryParseStyleDimension(string? style, string propertyName, out double value) {
        value = default;
        return TryParseImageDimension(TryGetStyleDeclarationValue(style, propertyName), out value);
    }

    private static string? TryGetStyleDeclarationValue(string? style, string propertyName) {
        if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName)) {
            return null;
        }

        foreach (var declaration in style!.Split(';')) {
            if (string.IsNullOrWhiteSpace(declaration)) {
                continue;
            }

            int separatorIndex = declaration.IndexOf(':');
            if (separatorIndex <= 0) {
                continue;
            }

            string name = declaration.Substring(0, separatorIndex).Trim();
            if (!name.Equals(propertyName, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string value = declaration.Substring(separatorIndex + 1).Trim();
            return value.Length == 0 ? null : value;
        }

        return null;
    }

    private static bool TryParseImageDimension(string? rawValue, out double value) {
        value = default;
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return false;
        }

        string normalized = rawValue!.Trim();
        if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) {
            normalized = normalized.Substring(0, normalized.Length - 2).Trim();
        }

        if (normalized.Length == 0
            || normalized.IndexOf('%') >= 0
            || normalized.IndexOf("calc(", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0) {
            return false;
        }

        return double.TryParse(
            normalized,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out value);
    }
}
