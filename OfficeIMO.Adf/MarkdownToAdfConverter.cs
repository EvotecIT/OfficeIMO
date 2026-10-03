using OfficeIMO.Markdown;

namespace OfficeIMO.Adf;

internal sealed class MarkdownToAdfConverter {
    private readonly Func<string, string>? _localIdFactory;
    private string _identityRoot = string.Empty;
    private readonly List<(AdfNode Node, string Path)> _taskIdentities = new List<(AdfNode, string)>();
    private readonly HashSet<string> _localIds = new HashSet<string>(StringComparer.Ordinal);

    private MarkdownToAdfConverter(AdfConversionOptions options) {
        _localIdFactory = options.LocalIdFactory;
    }

    internal static AdfDocument Convert(MarkdownDoc source, List<AdfConversionDiagnostic> diagnostics, AdfConversionOptions options) {
        var converter = new MarkdownToAdfConverter(options);
        AdfDocument document = converter.ConvertCore(source, diagnostics);
        converter.AssignTaskIdentities(document, options);
        return document;
    }

    private void AssignTaskIdentities(AdfDocument document, AdfConversionOptions options) {
        if (_taskIdentities.Count == 0) return;
        AdfGraphSafety.EnsureSafe(document, options);
        if (_localIdFactory == null) {
            // Hash bounded canonical content only when default task IDs are needed.
            // No recursive Markdown rendering occurs before resource checks.
            string json = document.ToJson(options);
            using var hash = System.Security.Cryptography.SHA256.Create();
            _identityRoot = BitConverter.ToString(hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(json)));
        }
        foreach (var task in _taskIdentities) {
            options.CancellationToken.ThrowIfCancellationRequested();
            task.Node.SetAttribute("localId", LocalId(task.Path));
        }
    }

    private AdfDocument ConvertCore(MarkdownDoc source, List<AdfConversionDiagnostic> diagnostics) {
        var document = new AdfDocument();
        for (int i = 0; i < source.Blocks.Count; i++) {
            AdfNode? node = ConvertBlock(source.Blocks[i], "$.blocks[" + i + "]", diagnostics);
            if (node != null) document.Content.Add(node);
        }
        return document;
    }

    private AdfNode? ConvertBlock(IMarkdownBlock block, string path, List<AdfConversionDiagnostic> diagnostics, string parentType = "doc") {
        switch (block) {
            case ImageBlock imageBlock:
                if (!string.IsNullOrEmpty(imageBlock.Title) || !string.IsNullOrEmpty(imageBlock.Caption) || imageBlock.Width.HasValue || imageBlock.Height.HasValue || !string.IsNullOrEmpty(imageBlock.LinkUrl) || imageBlock.PictureSources.Count > 0)
                    diagnostics.Add(Warning("MARKDOWN_IMAGE_PROPERTIES_DROPPED", path, "ADF image projection retains the source and alternate text; Markdown image title, caption, sizing, picture sources and wrapping link properties were omitted."));
                var blockMedia = new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", imageBlock.Path).SetAttribute("alt", imageBlock.PlainAlt ?? string.Empty);
                return new AdfNode("mediaSingle") { Content = { blockMedia } }.SetAttribute("layout", "center");
            case ParagraphBlock paragraph:
                if (paragraph.Inlines.Nodes.Count == 1 && paragraph.Inlines.Nodes[0] is ImageInline image) {
                    if (!string.IsNullOrEmpty(image.Title)) diagnostics.Add(Warning("MARKDOWN_IMAGE_TITLE_DROPPED", path, "ADF external media does not preserve Markdown image titles."));
                    var media = new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", image.Src).SetAttribute("alt", image.PlainAlt);
                    return new AdfNode("mediaSingle") { Content = { media } }.SetAttribute("layout", "center");
                }
                return WithInlines(new AdfNode("paragraph"), paragraph.Inlines, path, diagnostics);
            case HeadingBlock heading:
                return WithInlines(new AdfNode("heading").SetAttribute("level", heading.Level), heading.Inlines, path, diagnostics);
            case CodeBlock code:
                var codeNode = new AdfNode("codeBlock");
                if (!string.IsNullOrWhiteSpace(code.Language)) codeNode.SetAttribute("language", code.Language);
                if (!string.IsNullOrEmpty(code.Content)) codeNode.Content.Add(AdfNode.TextNode(code.Content));
                return codeNode;
            case QuoteBlock quote:
                var quoteNode = new AdfNode("blockquote");
                if (quote.ChildBlocks.Count > 0) {
                    AddBlocks(quoteNode, quote.ChildBlocks, path + ".children", diagnostics);
                } else {
                    foreach (string line in quote.Lines) quoteNode.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(line) } });
                }
                return quoteNode.Content.Count == 0 ? OmitEmptyBlock(path, diagnostics, "quote") : quoteNode;
            case UnorderedListBlock unordered:
                if (unordered.Items.Count == 0) return OmitEmptyBlock(path, diagnostics, "list");
                if (parentType != "blockquote" && CanConvertTaskList(unordered.Items)) return ConvertTaskList(unordered.Items, path, diagnostics);
                return ConvertList("bulletList", unordered.Items, path, diagnostics, 1);
            case OrderedListBlock ordered:
                if (ordered.Items.Count == 0) return OmitEmptyBlock(path, diagnostics, "list");
                return ConvertList("orderedList", ordered.Items, path, diagnostics, ordered.Start);
            case HorizontalRuleBlock:
                return new AdfNode("rule");
            case TableBlock table:
                if (table.HeaderInlines.Count == 0 && table.RowInlines.Count == 0) return OmitEmptyBlock(path, diagnostics, "table");
                return ConvertTable(table, path, diagnostics);
            case HtmlRawBlock raw:
                diagnostics.Add(Warning("MARKDOWN_RAW_HTML", path, "Raw HTML has no exact ADF mapping and was retained as an extension node."));
                return new AdfNode("extension")
                    .SetAttribute("extensionType", "com.officeimo.raw-html")
                    .SetAttribute("extensionKey", "raw-html")
                    .SetAttribute("parameters", new Dictionary<string, string> { ["html"] = raw.Html });
            default:
                diagnostics.Add(Warning("MARKDOWN_UNSUPPORTED_BLOCK", path, "Markdown block '" + block.GetType().Name + "' has no exact ADF mapping and was omitted."));
                return null;
        }
    }

    private AdfNode ConvertList(string type, IReadOnlyList<ListItem> items, string path, List<AdfConversionDiagnostic> diagnostics, int order) {
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

    private bool CanConvertTaskList(IReadOnlyList<ListItem> items) =>
        items.Count > 0 && items.All(item => item.IsTask && item.AdditionalParagraphs.Count == 0 &&
            item.NestedBlocks.All(block => block is UnorderedListBlock nested && CanConvertTaskList(nested.Items)));

    private AdfNode ConvertTaskList(IReadOnlyList<ListItem> items, string path, List<AdfConversionDiagnostic> diagnostics) {
        var list = new AdfNode("taskList");
        _taskIdentities.Add((list, path));
        for (int i = 0; i < items.Count; i++) {
            ListItem sourceItem = items[i];
            var item = new AdfNode("taskItem")
                .SetAttribute("state", sourceItem.Checked ? "DONE" : "TODO");
            _taskIdentities.Add((item, path + ".items[" + i + "]"));
            WithInlines(item, sourceItem.Content, path + ".items[" + i + "]", diagnostics);
            list.Content.Add(item);
            for (int nestedIndex = 0; nestedIndex < sourceItem.NestedBlocks.Count; nestedIndex++) {
                var nested = (UnorderedListBlock)sourceItem.NestedBlocks[nestedIndex];
                list.Content.Add(ConvertTaskList(nested.Items, path + ".items[" + i + "].nested[" + nestedIndex + "]", diagnostics));
            }
        }
        return list;
    }

    private string LocalId(string path) {
        string id;
        if (_localIdFactory != null) id = _localIdFactory(path);
        else {
            using var hash = System.Security.Cryptography.SHA256.Create();
            byte[] bytes = hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(_identityRoot + "\n" + path));
            var guidBytes = new byte[16];
            Array.Copy(bytes, guidBytes, guidBytes.Length);
            id = new Guid(guidBytes).ToString("D");
        }
        if (string.IsNullOrWhiteSpace(id) || !_localIds.Add(id)) throw new InvalidOperationException("ADF task localId values must be nonempty and unique within a conversion.");
        return id;
    }

    private AdfNode ConvertTable(TableBlock table, string path, List<AdfConversionDiagnostic> diagnostics) {
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

    private AdfNode WithInlines(AdfNode target, InlineSequence sequence, string path, List<AdfConversionDiagnostic> diagnostics) {
        AppendInlines(target.Content, sequence, Array.Empty<AdfMark>(), path, diagnostics);
        return target;
    }

    private void AppendInlines(List<AdfNode> target, InlineSequence sequence, IReadOnlyList<AdfMark> inheritedMarks, string path, List<AdfConversionDiagnostic> diagnostics) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            IMarkdownInline inline = sequence.Nodes[i];
            string inlinePath = path + ".inlines[" + i + "]";
            switch (inline) {
                case MarkdownTextRun text:
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
                case UnderlineInline underline:
                    target.Add(AdfNode.TextNode(underline.Text, AddMark(inheritedMarks, new AdfMark("underline"))));
                    break;
                case SuperscriptInline superscript:
                    target.Add(AdfNode.TextNode(superscript.Text, AddMark(inheritedMarks, ScriptMark("sup"))));
                    break;
                case SubscriptInline subscript:
                    target.Add(AdfNode.TextNode(subscript.Text, AddMark(inheritedMarks, ScriptMark("sub"))));
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
                case SuperscriptSequenceInline superscriptSequence:
                    AppendInlines(target, superscriptSequence.Inlines, AddMark(inheritedMarks, ScriptMark("sup")), inlinePath, diagnostics);
                    break;
                case SubscriptSequenceInline subscriptSequence:
                    AppendInlines(target, subscriptSequence.Inlines, AddMark(inheritedMarks, ScriptMark("sub")), inlinePath, diagnostics);
                    break;
                case HtmlTagSequenceInline htmlTag when htmlTag.TagName == "u":
                    AppendInlines(target, htmlTag.Inlines, AddMark(inheritedMarks, new AdfMark("underline")), inlinePath, diagnostics);
                    break;
                case HtmlTagSequenceInline htmlTag when htmlTag.TagName == "sup":
                    AppendInlines(target, htmlTag.Inlines, AddMark(inheritedMarks, ScriptMark("sup")), inlinePath, diagnostics);
                    break;
                case HtmlTagSequenceInline htmlTag when htmlTag.TagName == "sub":
                    AppendInlines(target, htmlTag.Inlines, AddMark(inheritedMarks, ScriptMark("sub")), inlinePath, diagnostics);
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
                case ImageInline image:
                    diagnostics.Add(Warning("MARKDOWN_INLINE_IMAGE_PROJECTED", inlinePath, "Inline Markdown images are represented by their linked alternate text; ADF external media requires a block container."));
                    target.Add(AdfNode.TextNode(image.PlainAlt.Length == 0 ? image.Src : image.PlainAlt,
                        inheritedMarks.Any(mark => mark.Type == "link") ? CloneMarks(inheritedMarks) :
                        AddMark(inheritedMarks, new AdfMark("link").SetAttribute("href", image.Src))));
                    break;
                case HtmlRawInline html when html.Html == "<!-- -->":
                    // Empty comments separate adjacent Markdown delimiters and carry no content.
                    break;
                default:
                    diagnostics.Add(Warning("MARKDOWN_UNSUPPORTED_INLINE", inlinePath, "Markdown inline '" + inline.GetType().Name + "' has no exact ADF mapping and was omitted."));
                    break;
            }
        }
    }

    private IReadOnlyList<AdfMark> AddMark(IReadOnlyList<AdfMark> marks, AdfMark added) {
        var result = CloneMarks(marks).ToList();
        result.Add(added);
        return result;
    }

    private AdfMark ScriptMark(string type) => new AdfMark("subsup").SetAttribute("type", type);

    private IReadOnlyList<AdfMark> CloneMarks(IReadOnlyList<AdfMark> marks) {
        var result = new List<AdfMark>(marks.Count);
        foreach (AdfMark source in marks) {
            var copy = new AdfMark(source.Type);
            foreach (var attribute in source.Attributes) copy.Attributes[attribute.Key] = attribute.Value.Clone();
            foreach (var extension in source.ExtensionData) copy.ExtensionData[extension.Key] = extension.Value.Clone();
            result.Add(copy);
        }
        return result;
    }

    private void AddBlocks(AdfNode target, IReadOnlyList<IMarkdownBlock> blocks, string path, List<AdfConversionDiagnostic> diagnostics) {
        for (int i = 0; i < blocks.Count; i++) {
            string childPath = path + "[" + i + "]";
            AdfNode? converted = ConvertBlock(blocks[i], childPath, diagnostics, target.Type);
            if (converted == null) continue;
            bool restricted = target.Type == "blockquote" || target.Type == "listItem";
            bool allowed = converted.Type == "paragraph" || converted.Type == "bulletList" || converted.Type == "orderedList" ||
                converted.Type == "codeBlock" || converted.Type == "mediaSingle" || converted.Type == "extension" ||
                target.Type == "blockquote" && converted.Type == "mediaGroup" || target.Type == "listItem" && converted.Type == "taskList";
            if (restricted && !allowed) {
                diagnostics.Add(Warning("MARKDOWN_CONTEXT_FALLBACK", childPath, "The Markdown block is not allowed in this ADF container and is preserved as visible Markdown source in a paragraph."));
                converted = new AdfNode("paragraph") { Content = { AdfNode.TextNode(blocks[i].RenderMarkdown()) } };
            }
            target.Content.Add(converted);
        }
    }

    private AdfConversionDiagnostic Warning(string code, string path, string message) => new AdfConversionDiagnostic(code, path, message, AdfConversionSeverity.Warning);

    private AdfNode? OmitEmptyBlock(string path, List<AdfConversionDiagnostic> diagnostics, string type) {
        diagnostics.Add(Warning("MARKDOWN_EMPTY_BLOCK_OMITTED", path, "An empty Markdown " + type + " cannot satisfy the ADF content contract and was omitted."));
        return null;
    }
}
