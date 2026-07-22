using OfficeIMO.Markdown;

namespace OfficeIMO.Adf;

internal static class MarkdownToAdfConverter {
    internal static AdfDocument Convert(MarkdownDoc source, List<AdfConversionDiagnostic> diagnostics) {
        var document = new AdfDocument();
        for (int i = 0; i < source.Blocks.Count; i++) {
            AdfNode? node = ConvertBlock(source.Blocks[i], "$.blocks[" + i + "]", diagnostics);
            if (node != null) document.Content.Add(node);
        }
        return document;
    }

    private static AdfNode? ConvertBlock(IMarkdownBlock block, string path, List<AdfConversionDiagnostic> diagnostics) {
        switch (block) {
            case ParagraphBlock paragraph:
                return WithInlines(new AdfNode("paragraph"), paragraph.Inlines, path, diagnostics);
            case HeadingBlock heading:
                return WithInlines(new AdfNode("heading").SetAttribute("level", heading.Level), heading.Inlines, path, diagnostics);
            case CodeBlock code:
                var codeNode = new AdfNode("codeBlock");
                if (!string.IsNullOrWhiteSpace(code.Language)) codeNode.SetAttribute("language", code.Language);
                codeNode.Content.Add(AdfNode.TextNode(code.Content));
                return codeNode;
            case QuoteBlock quote:
                var quoteNode = new AdfNode("blockquote");
                if (quote.ChildBlocks.Count > 0) {
                    AddBlocks(quoteNode, quote.ChildBlocks, path + ".children", diagnostics);
                } else {
                    foreach (string line in quote.Lines) quoteNode.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(line) } });
                }
                return quoteNode;
            case UnorderedListBlock unordered:
                if (CanConvertTaskList(unordered.Items)) return ConvertTaskList(unordered.Items, path, diagnostics);
                return ConvertList("bulletList", unordered.Items, path, diagnostics, 1);
            case OrderedListBlock ordered:
                return ConvertList("orderedList", ordered.Items, path, diagnostics, ordered.Start);
            case HorizontalRuleBlock:
                return new AdfNode("rule");
            case TableBlock table:
                return ConvertTable(table, path, diagnostics);
            case HtmlRawBlock raw:
                diagnostics.Add(Warning("MARKDOWN_RAW_HTML", path, "Raw HTML has no exact ADF mapping and was retained as an extension node."));
                return new AdfNode("extension")
                    .SetAttribute("extensionType", "com.officeimo.raw-html")
                    .SetAttribute("extensionKey", "raw-html")
                    .SetAttribute("parameters", new { html = raw.Html });
            default:
                diagnostics.Add(Warning("MARKDOWN_UNSUPPORTED_BLOCK", path, "Markdown block '" + block.GetType().Name + "' has no exact ADF mapping and was omitted."));
                return null;
        }
    }

    private static AdfNode ConvertList(string type, IReadOnlyList<ListItem> items, string path, List<AdfConversionDiagnostic> diagnostics, int order) {
        var list = new AdfNode(type);
        if (type == "orderedList" && order != 1) list.SetAttribute("order", order);
        for (int i = 0; i < items.Count; i++) {
            ListItem sourceItem = items[i];
            var item = new AdfNode("listItem");
            AdfNode listParagraph = WithInlines(new AdfNode("paragraph"), sourceItem.Content, path + ".items[" + i + "]", diagnostics);
            if (sourceItem.IsTask) {
                listParagraph.Content.Insert(0, AdfNode.TextNode(sourceItem.Checked ? "[x] " : "[ ] "));
                diagnostics.Add(Warning(
                    "MARKDOWN_TASK_LIST_FALLBACK",
                    path + ".items[" + i + "]",
                    "Markdown task state is preserved as a visible marker because ADF taskItem nodes require a taskList parent."));
            }
            item.Content.Add(listParagraph);
            foreach (InlineSequence paragraph in sourceItem.AdditionalParagraphs) {
                item.Content.Add(WithInlines(new AdfNode("paragraph"), paragraph, path + ".items[" + i + "]", diagnostics));
            }
            AddBlocks(item, sourceItem.NestedBlocks, path + ".items[" + i + "].nested", diagnostics);
            list.Content.Add(item);
        }
        return list;
    }

    private static bool CanConvertTaskList(IReadOnlyList<ListItem> items) =>
        items.Count > 0 && items.All(item => item.IsTask && item.AdditionalParagraphs.Count == 0 && item.NestedBlocks.Count == 0);

    private static AdfNode ConvertTaskList(IReadOnlyList<ListItem> items, string path, List<AdfConversionDiagnostic> diagnostics) {
        var list = new AdfNode("taskList").SetAttribute("localId", Guid.NewGuid().ToString("D"));
        for (int i = 0; i < items.Count; i++) {
            ListItem sourceItem = items[i];
            var item = new AdfNode("taskItem")
                .SetAttribute("localId", Guid.NewGuid().ToString("D"))
                .SetAttribute("state", sourceItem.Checked ? "DONE" : "TODO");
            WithInlines(item, sourceItem.Content, path + ".items[" + i + "]", diagnostics);
            list.Content.Add(item);
        }
        return list;
    }

    private static AdfNode ConvertTable(TableBlock table, string path, List<AdfConversionDiagnostic> diagnostics) {
        var result = new AdfNode("table");
        if (table.HeaderInlines.Count > 0) {
            var header = new AdfNode("tableRow");
            for (int column = 0; column < table.HeaderInlines.Count; column++) {
                var cell = new AdfNode("tableHeader");
                cell.Content.Add(WithInlines(new AdfNode("paragraph"), table.HeaderInlines[column], path + ".header[" + column + "]", diagnostics));
                header.Content.Add(cell);
            }
            result.Content.Add(header);
        }

        for (int row = 0; row < table.RowInlines.Count; row++) {
            var rowNode = new AdfNode("tableRow");
            for (int column = 0; column < table.RowInlines[row].Count; column++) {
                var cell = new AdfNode("tableCell");
                cell.Content.Add(WithInlines(new AdfNode("paragraph"), table.RowInlines[row][column], path + ".rows[" + row + "][" + column + "]", diagnostics));
                rowNode.Content.Add(cell);
            }
            result.Content.Add(rowNode);
        }
        return result;
    }

    private static AdfNode WithInlines(AdfNode target, InlineSequence sequence, string path, List<AdfConversionDiagnostic> diagnostics) {
        AppendInlines(target.Content, sequence, Array.Empty<AdfMark>(), path, diagnostics);
        return target;
    }

    private static void AppendInlines(List<AdfNode> target, InlineSequence sequence, IReadOnlyList<AdfMark> inheritedMarks, string path, List<AdfConversionDiagnostic> diagnostics) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            IMarkdownInline inline = sequence.Nodes[i];
            string inlinePath = path + ".inlines[" + i + "]";
            switch (inline) {
                case TextRun text:
                    target.Add(AdfNode.TextNode(text.Text, CloneMarks(inheritedMarks)));
                    break;
                case ILiteralTextMarkdownInline literalText:
                    target.Add(AdfNode.TextNode(literalText.Text, CloneMarks(inheritedMarks)));
                    break;
                case BoldInline bold:
                    target.Add(AdfNode.TextNode(bold.Text, AddMark(inheritedMarks, new AdfMark("strong"))));
                    break;
                case ItalicInline italic:
                    target.Add(AdfNode.TextNode(italic.Text, AddMark(inheritedMarks, new AdfMark("em"))));
                    break;
                case BoldItalicInline boldItalic:
                    target.Add(AdfNode.TextNode(boldItalic.Text, AddMark(AddMark(inheritedMarks, new AdfMark("strong")), new AdfMark("em"))));
                    break;
                case StrikethroughInline strike:
                    target.Add(AdfNode.TextNode(strike.Text, AddMark(inheritedMarks, new AdfMark("strike"))));
                    break;
                case CodeSpanInline code:
                    target.Add(AdfNode.TextNode(code.Text, AddMark(inheritedMarks, new AdfMark("code"))));
                    break;
                case BoldSequenceInline boldSequence:
                    AppendInlines(target, boldSequence.Inlines, AddMark(inheritedMarks, new AdfMark("strong")), inlinePath, diagnostics);
                    break;
                case ItalicSequenceInline italicSequence:
                    AppendInlines(target, italicSequence.Inlines, AddMark(inheritedMarks, new AdfMark("em")), inlinePath, diagnostics);
                    break;
                case BoldItalicSequenceInline boldItalicSequence:
                    AppendInlines(target, boldItalicSequence.Inlines, AddMark(AddMark(inheritedMarks, new AdfMark("strong")), new AdfMark("em")), inlinePath, diagnostics);
                    break;
                case StrikethroughSequenceInline strikeSequence:
                    AppendInlines(target, strikeSequence.Inlines, AddMark(inheritedMarks, new AdfMark("strike")), inlinePath, diagnostics);
                    break;
                case LinkInline link:
                    var linkMark = new AdfMark("link").SetAttribute("href", link.Url);
                    if (!string.IsNullOrWhiteSpace(link.Title)) linkMark.SetAttribute("title", link.Title);
                    if (link.LabelInlines != null) AppendInlines(target, link.LabelInlines, AddMark(inheritedMarks, linkMark), inlinePath, diagnostics);
                    else target.Add(AdfNode.TextNode(link.Text, AddMark(inheritedMarks, linkMark)));
                    break;
                case HardBreakInline:
                    target.Add(new AdfNode("hardBreak"));
                    break;
                case SoftBreakInline:
                    target.Add(AdfNode.TextNode("\n"));
                    break;
                default:
                    diagnostics.Add(Warning("MARKDOWN_UNSUPPORTED_INLINE", inlinePath, "Markdown inline '" + inline.GetType().Name + "' has no exact ADF mapping and was omitted."));
                    break;
            }
        }
    }

    private static IReadOnlyList<AdfMark> AddMark(IReadOnlyList<AdfMark> marks, AdfMark added) {
        var result = CloneMarks(marks).ToList();
        result.Add(added);
        return result;
    }

    private static IReadOnlyList<AdfMark> CloneMarks(IReadOnlyList<AdfMark> marks) {
        var result = new List<AdfMark>(marks.Count);
        foreach (AdfMark source in marks) {
            var copy = new AdfMark(source.Type);
            foreach (var attribute in source.Attributes) copy.Attributes[attribute.Key] = attribute.Value.Clone();
            foreach (var extension in source.ExtensionData) copy.ExtensionData[extension.Key] = extension.Value.Clone();
            result.Add(copy);
        }
        return result;
    }

    private static void AddBlocks(AdfNode target, IReadOnlyList<IMarkdownBlock> blocks, string path, List<AdfConversionDiagnostic> diagnostics) {
        for (int i = 0; i < blocks.Count; i++) {
            AdfNode? converted = ConvertBlock(blocks[i], path + "[" + i + "]", diagnostics);
            if (converted != null) target.Content.Add(converted);
        }
    }

    private static AdfConversionDiagnostic Warning(string code, string path, string message) => new AdfConversionDiagnostic(code, path, message, AdfConversionSeverity.Warning);
}
