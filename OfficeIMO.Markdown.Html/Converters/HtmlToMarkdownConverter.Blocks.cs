using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static readonly HashSet<string> s_BlockTags = new(StringComparer.OrdinalIgnoreCase) {
        "ADDRESS", "ARTICLE", "ASIDE", "BLOCKQUOTE", "BODY", "DETAILS", "DIV", "DL", "FIGURE",
        "FOOTER", "FORM", "H1", "H2", "H3", "H4", "H5", "H6", "HEADER", "HR", "LI", "MAIN",
        "NAV", "OL", "P", "PRE", "SECTION", "TABLE", "UL"
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

            string inlineMarkdown = ConvertInlineNodesToMarkdown(inlineBuffer, context).Trim();
            inlineBuffer.Clear();
            if (inlineMarkdown.Length == 0) {
                return;
            }

            blocks.Add(new ParagraphBlock(ParseInlines(inlineMarkdown)));
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

        if (HasDirectBlockChildren(element, context)) {
            return true;
        }

        if (context.Options.PreserveUnsupportedBlocks && !IsInlineElement(element)) {
            return true;
        }

        return false;
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
        if (TryConvertVisualContractElement(element, out var visualBlock)) {
            return new IMarkdownBlock[] { visualBlock };
        }

        if (TryConvertMermaidElement(element, out var mermaidBlock)) {
            return new IMarkdownBlock[] { mermaidBlock };
        }

        if (TryConvertMathElement(element, out var mathBlock)) {
            return new IMarkdownBlock[] { mathBlock };
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
            case "FIGURE":
                return ConvertFigureElement(element, context);
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

                string inlineMarkdown = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
                if (inlineMarkdown.Length == 0) {
                    return Array.Empty<IMarkdownBlock>();
                }

                return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(inlineMarkdown)) };
            default:
                if (context.Options.PreserveUnsupportedBlocks) {
                    return new IMarkdownBlock[] { new HtmlRawBlock(element.OuterHtml) };
                }

                if (HasDirectBlockChildren(element, context)) {
                    return ConvertNodesToBlocks(element.ChildNodes, context);
                }

                string fallbackInline = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
                if (fallbackInline.Length == 0) {
                    return Array.Empty<IMarkdownBlock>();
                }

                return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(fallbackInline)) };
        }
    }

    private static bool TryConvertVisualContractElement(IElement element, out SemanticFencedBlock visualBlock) {
        visualBlock = null!;
        if (element == null) {
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

        visualBlock = new SemanticFencedBlock(visualElement.VisualKind, visualElement.FenceLanguage, payload);
        return true;
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
        string inlineMarkdown = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
        if (inlineMarkdown.Length == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(inlineMarkdown)) };
    }

    private static HeadingBlock ConvertHeadingElement(IElement element, int level, ConversionContext context) {
        string inlineMarkdown = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
        return new HeadingBlock(level, ParseInlines(inlineMarkdown));
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
        for (int i = 0; i < element.ClassList.Length; i++) {
            string token = element.ClassList[i];
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
        if (blocks[0] is ParagraphBlock firstParagraph
            && TryExtractCalloutTitleFromParagraph(firstParagraph, out var extractedTitle)) {
            titleInlines = extractedTitle;
            blocks.RemoveAt(0);
        }

        callout = new CalloutBlock(kind, titleInlines, blocks);
        return true;
    }

    private static bool TryExtractCalloutTitleFromParagraph(ParagraphBlock paragraph, out InlineSequence titleInlines) {
        titleInlines = new InlineSequence();
        if (paragraph == null) {
            return false;
        }

        string markdown = paragraph.Inlines.RenderMarkdown().Trim();
        if (markdown.Length < 4
            || !markdown.StartsWith("**", StringComparison.Ordinal)
            || !markdown.EndsWith("**", StringComparison.Ordinal)) {
            return false;
        }

        string inner = markdown.Substring(2, markdown.Length - 4).Trim();
        if (inner.Length == 0) {
            return false;
        }

        titleInlines = ParseInlines(inner);
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

        foreach (var row in element.QuerySelectorAll("tr")) {
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

    private static ColumnAlignment ParseAlignment(IElement cell) {
        string? align = cell.GetAttribute("align");
        if (string.IsNullOrWhiteSpace(align)) {
            return ColumnAlignment.None;
        }

        switch (align!.Trim().ToLowerInvariant()) {
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

        string inlineMarkdown = ConvertInlineNodesToMarkdown(cell.ChildNodes, context).Trim();
        if (inlineMarkdown.Length == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(inlineMarkdown)) };
    }

    private static string RenderTableCellBlocksToMarkdown(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        return new TableCell(blocks).Markdown.Replace("  \n", "<br>");
    }

    private static IEnumerable<IMarkdownBlock> ConvertImageElement(IElement element, ConversionContext context) {
        string src = ResolveUrl(element.GetAttribute("src"), context);
        if (string.IsNullOrWhiteSpace(src)) {
            return Array.Empty<IMarkdownBlock>();
        }

        var image = new ImageBlock(
            src,
            element.GetAttribute("alt"),
            element.GetAttribute("title"));
        if (double.TryParse(element.GetAttribute("width"), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double width)) {
            image.Width = width;
        }
        if (double.TryParse(element.GetAttribute("height"), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double height)) {
            image.Height = height;
        }

        return new IMarkdownBlock[] { image };
    }

    private static IEnumerable<IMarkdownBlock> ConvertFigureElement(IElement element, ConversionContext context) {
        var image = element.QuerySelector("img");
        if (image == null) {
            if (HasDirectBlockChildren(element, context)) {
                return ConvertNodesToBlocks(element.ChildNodes, context);
            }

            string inlineMarkdown = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
            if (inlineMarkdown.Length == 0) {
                return Array.Empty<IMarkdownBlock>();
            }

            return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(inlineMarkdown)) };
        }

        var blocks = ConvertImageElement(image, context).ToList();
        var caption = element.QuerySelector("figcaption");
        if (caption != null && blocks.Count == 1 && blocks[0] is ImageBlock imageBlock) {
            imageBlock.Caption = NormalizeBlockText(caption.TextContent);
        }

        return blocks;
    }

    private static DetailsBlock ConvertDetailsElement(IElement element, ConversionContext context) {
        SummaryBlock? summary = null;
        var summaryElement = element.Children.FirstOrDefault(child => child.TagName.Equals("SUMMARY", StringComparison.OrdinalIgnoreCase));
        if (summaryElement != null) {
            summary = new SummaryBlock(ParseInlines(ConvertInlineNodesToMarkdown(summaryElement.ChildNodes, context).Trim()));
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
        var pendingTerms = new List<string>();
        bool hasDefinitionsForCurrentGroup = false;

        foreach (var child in element.Children) {
            if (child.TagName.Equals("DT", StringComparison.OrdinalIgnoreCase)) {
                if (hasDefinitionsForCurrentGroup) {
                    pendingTerms.Clear();
                    hasDefinitionsForCurrentGroup = false;
                }

                string term = ConvertInlineNodesToMarkdown(child.ChildNodes, context).Trim();
                if (term.Length > 0) {
                    pendingTerms.Add(term);
                }
                continue;
            }

            if (child.TagName.Equals("DD", StringComparison.OrdinalIgnoreCase) && pendingTerms.Count > 0) {
                foreach (string term in pendingTerms) {
                    list.AddEntry(new DefinitionListEntry(
                        ParseInlines(term),
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

        string inlineMarkdown = ConvertInlineNodesToMarkdown(element.ChildNodes, context).Trim();
        if (inlineMarkdown.Length == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(ParseInlines(inlineMarkdown)) };
    }
}
