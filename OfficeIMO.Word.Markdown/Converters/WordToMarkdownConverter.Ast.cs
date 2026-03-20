using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OmdListItem = OfficeIMO.Markdown.ListItem;
using OmdTableCell = OfficeIMO.Markdown.TableCell;

namespace OfficeIMO.Word.Markdown {
    internal partial class WordToMarkdownConverter {
        private sealed class PendingListFrame {
            public PendingListFrame(int level, bool ordered, IMarkdownListBlock block) {
                Level = level;
                Ordered = ordered;
                Block = block;
            }

            public int Level { get; }
            public bool Ordered { get; }
            public IMarkdownListBlock Block { get; }
            public OmdListItem? LastItem { get; private set; }

            public void AddItem(OmdListItem item) {
                LastItem = item;
                switch (Block) {
                    case OrderedListBlock ordered:
                        ordered.Items.Add(item);
                        break;
                    case UnorderedListBlock unordered:
                        unordered.Items.Add(item);
                        break;
                }
            }
        }

        private void BuildMarkdownDocument(WordDocument document, MarkdownDoc markdown, WordToMarkdownOptions options, CancellationToken cancellationToken) {
            int sectionIndex = 0;
            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (options.IncludeHeadersAndFootersAsSemanticBlocks) {
                    AppendHeaderFooterSemanticBlocks(markdown, section, options, cancellationToken, sectionIndex);
                }

                var elements = section.Elements;
                if (elements == null || elements.Count == 0) {
                    elements = new List<WordElement>(section.Paragraphs.Count + section.Tables.Count);
                    elements.AddRange(section.Paragraphs);
                    elements.AddRange(section.Tables);
                }

                AppendBlocksFromElements(
                    elements,
                    block => markdown.Add(block),
                    options,
                    cancellationToken,
                    allowQuoteHeuristic: true,
                    trimBoundaryWhitespace: false);

                if (options.IncludeHeadersAndFootersAsSemanticBlocks) {
                    AppendFooterSemanticBlocks(markdown, section, options, cancellationToken, sectionIndex);
                }

                sectionIndex++;
            }

            AppendFootnotes(document, markdown, options);
        }

        private void AppendHeaderFooterSemanticBlocks(
            MarkdownDoc markdown,
            WordSection section,
            WordToMarkdownOptions options,
            CancellationToken cancellationToken,
            int sectionIndex) {
            AppendHeaderFooterSemanticBlock(markdown, section.Header.Default, options, cancellationToken, sectionIndex, "default", isHeader: true);
            AppendHeaderFooterSemanticBlock(markdown, section.Header.First, options, cancellationToken, sectionIndex, "first", isHeader: true);
            AppendHeaderFooterSemanticBlock(markdown, section.Header.Even, options, cancellationToken, sectionIndex, "even", isHeader: true);
        }

        private void AppendFooterSemanticBlocks(
            MarkdownDoc markdown,
            WordSection section,
            WordToMarkdownOptions options,
            CancellationToken cancellationToken,
            int sectionIndex) {
            AppendHeaderFooterSemanticBlock(markdown, section.Footer.Default, options, cancellationToken, sectionIndex, "default", isHeader: false);
            AppendHeaderFooterSemanticBlock(markdown, section.Footer.First, options, cancellationToken, sectionIndex, "first", isHeader: false);
            AppendHeaderFooterSemanticBlock(markdown, section.Footer.Even, options, cancellationToken, sectionIndex, "even", isHeader: false);
        }

        private void AppendHeaderFooterSemanticBlock(
            MarkdownDoc markdown,
            WordHeaderFooter? headerFooter,
            WordToMarkdownOptions options,
            CancellationToken cancellationToken,
            int sectionIndex,
            string slot,
            bool isHeader) {
            if (headerFooter == null) {
                return;
            }

            var blocks = new List<IMarkdownBlock>();
            AppendBlocksFromElements(
                headerFooter.Elements,
                block => blocks.Add(block),
                options,
                cancellationToken,
                allowQuoteHeuristic: true,
                trimBoundaryWhitespace: false);

            if (blocks.Count == 0) {
                return;
            }

            var infoString = BuildHeaderFooterFenceInfoString(isHeader, sectionIndex + 1, slot);
            var semanticKind = isHeader
                ? WordMarkdownSemanticBlocks.HeaderSemanticKind
                : WordMarkdownSemanticBlocks.FooterSemanticKind;

            markdown.Add(new SemanticFencedBlock(
                semanticKind,
                infoString,
                RenderMarkdownFragment(blocks)));
        }

        private static string BuildHeaderFooterFenceInfoString(bool isHeader, int sectionNumber, string slot) {
            var language = isHeader
                ? WordMarkdownSemanticBlocks.HeaderFenceLanguage
                : WordMarkdownSemanticBlocks.FooterFenceLanguage;
            return $"{language} section={sectionNumber} slot={slot}";
        }

        private static string RenderMarkdownFragment(IReadOnlyList<IMarkdownBlock> blocks) {
            var fragment = MarkdownDoc.Create();
            for (int i = 0; i < blocks.Count; i++) {
                fragment.Add(blocks[i]);
            }

            return NormalizeMarkdownLineEndings(fragment.ToMarkdown());
        }

        private void AddListParagraph(
            Action<IMarkdownBlock> addRootBlock,
            List<PendingListFrame> listStack,
            WordParagraph paragraph,
            DocumentTraversal.ListInfo listInfo,
            WordToMarkdownOptions options,
            bool hasCheckbox,
            bool checkboxChecked,
            bool trimBoundaryWhitespace) {
            EnsureListFrame(addRootBlock, listStack, listInfo);

            var paragraphBlocks = BuildParagraphBlocks(paragraph, options, hasCheckbox, checkboxChecked, allowQuoteHeuristic: false, trimBoundaryWhitespace: trimBoundaryWhitespace);
            var item = CreateListItem(paragraphBlocks, listInfo.Level, hasCheckbox, checkboxChecked);
            listStack[listStack.Count - 1].AddItem(item);
        }

        private static void EnsureListFrame(Action<IMarkdownBlock> addRootBlock, List<PendingListFrame> listStack, DocumentTraversal.ListInfo listInfo) {
            int targetDepth = Math.Max(0, listInfo.Level) + 1;

            while (listStack.Count > targetDepth) {
                listStack.RemoveAt(listStack.Count - 1);
            }

            if (listStack.Count == targetDepth && listStack[targetDepth - 1].Ordered != listInfo.Ordered) {
                listStack.RemoveRange(targetDepth - 1, listStack.Count - (targetDepth - 1));
            }

            while (listStack.Count < targetDepth) {
                bool ordered = listInfo.Ordered;
                IMarkdownListBlock block = ordered ? new OrderedListBlock() : new UnorderedListBlock();
                if (block is OrderedListBlock orderedList && listStack.Count == targetDepth - 1) {
                    orderedList.Start = listInfo.Start;
                }

                if (listStack.Count == 0) {
                    addRootBlock((IMarkdownBlock)block);
                } else {
                    var parentFrame = listStack[listStack.Count - 1];
                    if (parentFrame.LastItem == null) {
                        var placeholder = new OmdListItem(new InlineSequence());
                        placeholder.Level = parentFrame.Level;
                        parentFrame.AddItem(placeholder);
                    }

                    parentFrame.LastItem!.Children.Add((IMarkdownBlock)block);
                }

                listStack.Add(new PendingListFrame(listStack.Count, ordered, block));
            }
        }

        private static OmdListItem CreateListItem(
            IReadOnlyList<IMarkdownBlock> paragraphBlocks,
            int level,
            bool hasCheckbox,
            bool checkboxChecked) {
            OmdListItem item;
            if (paragraphBlocks.Count > 0 && paragraphBlocks[0] is ParagraphBlock paragraphBlock) {
                item = hasCheckbox
                    ? OmdListItem.TaskInlines(paragraphBlock.Inlines, checkboxChecked)
                    : new OmdListItem(paragraphBlock.Inlines);

                for (int i = 1; i < paragraphBlocks.Count; i++) {
                    if (paragraphBlocks[i] is ParagraphBlock additionalParagraph) {
                        item.AdditionalParagraphs.Add(additionalParagraph.Inlines);
                    } else {
                        item.Children.Add(paragraphBlocks[i]);
                    }
                }
            } else {
                item = hasCheckbox
                    ? OmdListItem.TaskInlines(new InlineSequence(), checkboxChecked)
                    : new OmdListItem(new InlineSequence());

                for (int i = 0; i < paragraphBlocks.Count; i++) {
                    item.Children.Add(paragraphBlocks[i]);
                }
            }

            item.Level = Math.Max(0, level);
            return item;
        }

        private void AppendBlocksFromElements(
            IReadOnlyList<WordElement> elements,
            Action<IMarkdownBlock> addRootBlock,
            WordToMarkdownOptions options,
            CancellationToken cancellationToken,
            bool allowQuoteHeuristic,
            bool trimBoundaryWhitespace) {
            var listStack = new List<PendingListFrame>();

            for (int i = 0; i < elements.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                var element = elements[i];

                if (element is WordParagraph paragraph) {
                    if (paragraph.IsTextBox && paragraph.TextBox != null) {
                        listStack.Clear();
                        AppendBlocksFromElements(
                            paragraph.TextBox.Elements,
                            addRootBlock,
                            options,
                            cancellationToken,
                            allowQuoteHeuristic: allowQuoteHeuristic,
                            trimBoundaryWhitespace: true);
                        continue;
                    }

                    bool hasRuns = false;
                    try {
                        hasRuns = paragraph.GetRuns().Any();
                    } catch (InvalidOperationException ex) {
                        Debug.WriteLine($"GetRuns() failed for paragraph during Markdown AST conversion: {ex.Message}");
                        hasRuns = false;
                    }

                    ResolveParagraphCheckboxState(paragraph, out bool hasCheckbox, out bool checkboxChecked);
                    int backscan = i - 1;
                    while (backscan >= 0 && elements[backscan] is WordParagraph previous && previous.Equals(paragraph)) {
                        ResolveParagraphCheckboxState(previous, out bool previousHasCheckbox, out bool previousCheckboxChecked);
                        if (previousHasCheckbox) {
                            hasCheckbox = true;
                            checkboxChecked = previousCheckboxChecked;
                        }
                        backscan--;
                    }

                    int scan = i + 1;
                    while (scan < elements.Count && elements[scan] is WordParagraph sibling && sibling.Equals(paragraph)) {
                        ResolveParagraphCheckboxState(sibling, out bool siblingHasCheckbox, out bool siblingCheckboxChecked);
                        if (siblingHasCheckbox) {
                            hasCheckbox = true;
                            checkboxChecked = siblingCheckboxChecked;
                        }
                        scan++;
                    }

                    if (hasRuns && !paragraph.IsFirstRun) {
                        continue;
                    }

                    var listInfo = DocumentTraversal.GetListInfo(paragraph);
                    if (listInfo != null) {
                        AddListParagraph(addRootBlock, listStack, paragraph, listInfo.Value, options, hasCheckbox, checkboxChecked, trimBoundaryWhitespace);
                        continue;
                    }

                    listStack.Clear();
                    foreach (var block in BuildParagraphBlocks(paragraph, options, hasCheckbox, checkboxChecked, allowQuoteHeuristic, trimBoundaryWhitespace)) {
                        addRootBlock(block);
                    }
                    continue;
                }

                listStack.Clear();

                if (element is WordTable table) {
                    addRootBlock(BuildTableBlock(table, options));
                    continue;
                }

                if (element is WordEmbeddedDocument embeddedDocument) {
                    var html = embeddedDocument.GetHtml();
                    if (!string.IsNullOrWhiteSpace(html)) {
                        addRootBlock(new HtmlRawBlock(html!.TrimEnd()));
                    }
                }
            }
        }

        private IReadOnlyList<IMarkdownBlock> BuildParagraphBlocks(
            WordParagraph paragraph,
            WordToMarkdownOptions options,
            bool hasCheckbox,
            bool checkboxChecked,
            bool allowQuoteHeuristic,
            bool trimBoundaryWhitespace = false) {
            const string codeLangPrefix = "CodeLang_";
            var blocks = new List<IMarkdownBlock>();

            string? styleId = paragraph.StyleId;
            if (styleId is { Length: > 0 } sid && sid.StartsWith(codeLangPrefix, StringComparison.Ordinal)) {
                var runs = paragraph.GetRuns()
                    .Where(run => !string.IsNullOrEmpty(run.Text))
                    .ToList();

                if (runs.Count > 0) {
                    string language = sid.Substring(codeLangPrefix.Length);
                    string code = string.Concat(runs.Select(run => run.Text));
                    blocks.Add(new CodeBlock(language, code));
                    return blocks;
                }
            }

            var inlines = BuildParagraphInlines(paragraph, options, trimBoundaryWhitespace);
            if (inlines.Nodes.Count == 0 && !hasCheckbox) {
                return blocks;
            }

            int headingLevel = paragraph.Style.HasValue
                ? HeadingStyleMapper.GetLevelForHeadingStyle(paragraph.Style.Value)
                : 0;

            IMarkdownBlock block = headingLevel > 0
                ? new HeadingBlock(headingLevel, inlines)
                : new ParagraphBlock(inlines);

            if (allowQuoteHeuristic && paragraph.IndentationBefore.HasValue && paragraph.IndentationBefore.Value > 0) {
                int depth = (int)Math.Round(paragraph.IndentationBefore.Value / 720d);
                if (depth > 0) {
                    block = WrapQuotedBlock(block, depth);
                }
            }

            blocks.Add(block);
            return blocks;
        }

        private InlineSequence BuildParagraphInlines(WordParagraph paragraph, WordToMarkdownOptions options, bool trimBoundaryWhitespace = false) {
            var sequence = new InlineSequence { AutoSpacing = false };
            string? preferredCodeFont = ResolveConfiguredCodeFont(options.FontFamily);
            string? implicitCodeFont = ResolveImplicitCodeFont();

            foreach (var run in paragraph.GetRuns()) {
                if (run.Break != null && run.PageBreak == null) {
                    sequence.HardBreak();
                }

                if (run.IsFootNote && run.FootNote != null && run.FootNote.ReferenceId.HasValue) {
                    sequence.FootnoteRef(run.FootNote.ReferenceId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    continue;
                }

                if (run.IsImage && run.Image != null) {
                    sequence.AddRaw(CreateImageInline(run.Image, options));
                    continue;
                }

                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                AppendFormattedTextRun(sequence, run, text, options, preferredCodeFont, implicitCodeFont);
            }

            if (trimBoundaryWhitespace) {
                TrimBoundaryWhitespace(sequence);
            }

            return sequence;
        }

        private static void ResolveParagraphCheckboxState(WordParagraph paragraph, out bool hasCheckbox, out bool checkboxChecked) {
            hasCheckbox = paragraph.IsCheckBox;
            checkboxChecked = paragraph.CheckBox?.IsChecked == true;

            if (hasCheckbox) {
                return;
            }

            try {
                foreach (var run in paragraph.GetRuns()) {
                    if (!run.IsCheckBox) {
                        continue;
                    }

                    hasCheckbox = true;
                    checkboxChecked = run.CheckBox?.IsChecked == true;
                    return;
                }
            } catch (InvalidOperationException ex) {
                Debug.WriteLine($"GetRuns() failed while resolving checkbox state for Markdown AST conversion: {ex.Message}");
            }
        }

        private static void TrimBoundaryWhitespace(InlineSequence sequence) {
            if (sequence.Nodes.Count == 0) {
                return;
            }

            var nodes = sequence.Nodes.ToList();

            while (nodes.Count > 0 && nodes[0] is HardBreakInline) {
                nodes.RemoveAt(0);
            }

            while (nodes.Count > 0 && nodes[nodes.Count - 1] is HardBreakInline) {
                nodes.RemoveAt(nodes.Count - 1);
            }

            if (nodes.Count > 0 && nodes[0] is TextRun leadingText) {
                string trimmed = leadingText.Text.TrimStart();
                if (trimmed.Length == 0) {
                    nodes.RemoveAt(0);
                } else if (!string.Equals(trimmed, leadingText.Text, StringComparison.Ordinal)) {
                    nodes[0] = new TextRun(trimmed);
                }
            }

            if (nodes.Count > 0 && nodes[nodes.Count - 1] is TextRun trailingText) {
                string trimmed = trailingText.Text.TrimEnd();
                if (trimmed.Length == 0) {
                    nodes.RemoveAt(nodes.Count - 1);
                } else if (!string.Equals(trimmed, trailingText.Text, StringComparison.Ordinal)) {
                    nodes[nodes.Count - 1] = new TextRun(trimmed);
                }
            }

            sequence.ReplaceItems(nodes);
        }

        private void AppendFormattedTextRun(
            InlineSequence target,
            WordParagraph run,
            string text,
            WordToMarkdownOptions options,
            string? preferredCodeFont,
            string? implicitCodeFont) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            bool hasSemanticFormatting = RunHasSemanticFormatting(run, options, preferredCodeFont, implicitCodeFont);
            if (!hasSemanticFormatting) {
                target.Text(text);
                return;
            }

            SplitBoundaryWhitespace(text, out var leadingWhitespace, out var coreText, out var trailingWhitespace);
            if (!string.IsNullOrEmpty(leadingWhitespace)) {
                target.Text(leadingWhitespace);
            }

            if (!string.IsNullOrEmpty(coreText)) {
                var formatted = BuildFormattedRunInline(run, coreText, options, preferredCodeFont, implicitCodeFont);
                if (run.IsHyperLink && run.Hyperlink?.Uri != null) {
                    var label = new InlineSequence { AutoSpacing = false }.AddRaw(formatted);
                    string url = BuildHyperlinkUrl(run.Hyperlink.Uri);
                    formatted = new LinkInline(label, url, title: null);
                }

                target.AddRaw(formatted);
            }

            if (!string.IsNullOrEmpty(trailingWhitespace)) {
                target.Text(trailingWhitespace);
            }
        }

        private static bool RunHasSemanticFormatting(
            WordParagraph run,
            WordToMarkdownOptions options,
            string? preferredCodeFont,
            string? implicitCodeFont) {
            return run.Bold
                || run.Italic
                || run.Strike
                || run.VerticalTextAlignment == VerticalPositionValues.Superscript
                || run.VerticalTextAlignment == VerticalPositionValues.Subscript
                || (options.EnableUnderline && run.Underline.HasValue && run.Underline.Value != UnderlineValues.None)
                || (options.EnableHighlight && run.Highlight.HasValue && run.Highlight.Value != HighlightColorValues.None)
                || IsCodeRun(run, preferredCodeFont, implicitCodeFont)
                || (run.IsHyperLink && run.Hyperlink?.Uri != null);
        }

        private static IMarkdownInline BuildFormattedRunInline(
            WordParagraph run,
            string text,
            WordToMarkdownOptions options,
            string? preferredCodeFont,
            string? implicitCodeFont) {
            IMarkdownInline node = IsCodeRun(run, preferredCodeFont, implicitCodeFont)
                ? new CodeSpanInline(text)
                : new TextRun(text);

            if (run.VerticalTextAlignment == VerticalPositionValues.Superscript) {
                node = WrapInline("sup", node);
            } else if (run.VerticalTextAlignment == VerticalPositionValues.Subscript) {
                node = WrapInline("sub", node);
            }

            if (options.EnableUnderline && run.Underline.HasValue && run.Underline.Value != UnderlineValues.None) {
                node = WrapInline("u", node);
            }

            if (run.Strike) {
                node = new StrikethroughSequenceInline(WrapInlineSequence(node));
            }

            if (options.EnableHighlight && run.Highlight.HasValue && run.Highlight.Value != HighlightColorValues.None) {
                node = new HighlightSequenceInline(WrapInlineSequence(node));
            }

            if (run.Bold && run.Italic) {
                node = new BoldItalicSequenceInline(WrapInlineSequence(node));
            } else if (run.Bold) {
                node = new BoldSequenceInline(WrapInlineSequence(node));
            } else if (run.Italic) {
                node = new ItalicSequenceInline(WrapInlineSequence(node));
            }

            return node;
        }

        private static IMarkdownInline WrapInline(string tagName, IMarkdownInline node) =>
            new HtmlTagSequenceInline(tagName, WrapInlineSequence(node));

        private static InlineSequence WrapInlineSequence(IMarkdownInline node) {
            var sequence = new InlineSequence { AutoSpacing = false };
            sequence.AddRaw(node);
            return sequence;
        }

        private static void SplitBoundaryWhitespace(string text, out string leadingWhitespace, out string coreText, out string trailingWhitespace) {
            int start = 0;
            int end = text.Length - 1;

            while (start <= end && char.IsWhiteSpace(text[start])) {
                start++;
            }

            while (end >= start && char.IsWhiteSpace(text[end])) {
                end--;
            }

            if (start > end) {
                leadingWhitespace = text;
                coreText = string.Empty;
                trailingWhitespace = string.Empty;
                return;
            }

            leadingWhitespace = start == 0 ? string.Empty : text.Substring(0, start);
            coreText = text.Substring(start, end - start + 1);
            trailingWhitespace = end == text.Length - 1 ? string.Empty : text.Substring(end + 1);
        }

        private static string BuildHyperlinkUrl(Uri uri) {
            if (uri.IsAbsoluteUri) {
                string url = uri.GetComponents(UriComponents.AbsoluteUri, UriFormat.UriEscaped);
                var original = uri.OriginalString;
                if (!string.IsNullOrEmpty(original)
                    && !original.EndsWith("/", StringComparison.Ordinal)
                    && uri.AbsolutePath == "/"
                    && url.EndsWith("/", StringComparison.Ordinal)) {
                    url = url.TrimEnd('/');
                }

                return url;
            }

            return uri.ToString();
        }

        private static bool IsCodeRun(WordParagraph run, string? preferredCodeFont, string? implicitCodeFont) {
            string? runFont = run.FontFamily;
            if (string.IsNullOrEmpty(runFont)) {
                return false;
            }

            if (!string.IsNullOrEmpty(preferredCodeFont) && string.Equals(runFont, preferredCodeFont, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (!string.IsNullOrEmpty(implicitCodeFont) && string.Equals(runFont, implicitCodeFont, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string resolvedRunFont = runFont ?? string.Empty;
            if (resolvedRunFont.Length == 0) {
                return false;
            }

            return KnownMonospaceFonts.Contains(resolvedRunFont)
                || resolvedRunFont.IndexOf("Mono", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private ImageInline CreateImageInline(WordImage image, WordToMarkdownOptions options) {
            if (image == null) {
                return new ImageInline(string.Empty, string.Empty);
            }

            string alt = image.Description ?? string.Empty;
            string title = image.Title ?? string.Empty;
            string source = BuildImageSource(image, options);
            return new ImageInline(alt, source, string.IsNullOrEmpty(title) ? null : title);
        }

        private string BuildImageSource(WordImage image, WordToMarkdownOptions options) {
            if (options.ImageExportMode == ImageExportMode.File) {
                string directory = options.ImageDirectory ?? Directory.GetCurrentDirectory();
                Directory.CreateDirectory(directory);
                string fileExtension = Path.GetExtension(image.FilePath);
                if (string.IsNullOrEmpty(fileExtension)) {
                    fileExtension = ".png";
                }

                string fileName = string.IsNullOrEmpty(image.FileName)
                    ? Guid.NewGuid().ToString("N") + fileExtension
                    : image.FileName!;
                string targetPath = Path.Combine(directory, fileName);

                if (!string.IsNullOrEmpty(image.FilePath) && File.Exists(image.FilePath)) {
                    File.Copy(image.FilePath, targetPath, true);
                } else {
                    File.WriteAllBytes(targetPath, image.GetBytes());
                }

                return fileName;
            }

            byte[] bytes = image.GetBytes();
            string imageExtension = Path.GetExtension(image.FilePath);
            string mime = imageExtension switch {
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                _ => "image/png"
            };
            string base64 = System.Convert.ToBase64String(bytes);
            return $"data:{mime};base64,{base64}";
        }

        private TableBlock BuildTableBlock(WordTable table, WordToMarkdownOptions options) {
            var markdownTable = new TableBlock();
            var structuredHeaders = new List<OmdTableCell>();
            var structuredRows = new List<IReadOnlyList<OmdTableCell>>();

            if (table.Rows.Count == 0) {
                return markdownTable;
            }

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                var row = table.Rows[rowIndex];
                var structuredCells = new List<OmdTableCell>(row.Cells.Count);
                var markdownCells = new List<string>(row.Cells.Count);

                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                    var cell = row.Cells[cellIndex];
                    var cellBlock = BuildTableCell(cell, options);
                    structuredCells.Add(cellBlock);
                    markdownCells.Add(cellBlock.Markdown);

                    if (rowIndex == 0) {
                        markdownTable.Alignments.Add(GetAlignment(cell));
                    }
                }

                if (rowIndex == 0) {
                    for (int i = 0; i < markdownCells.Count; i++) {
                        markdownTable.Headers.Add(markdownCells[i]);
                    }
                    structuredHeaders.AddRange(structuredCells);
                } else {
                    markdownTable.Rows.Add(markdownCells);
                    structuredRows.Add(structuredCells);
                }
            }

            markdownTable.SetStructuredCells(structuredHeaders, structuredRows, markdownTable.ComputeContentSignature());
            return markdownTable;
        }

        private OmdTableCell BuildTableCell(WordTableCell cell, WordToMarkdownOptions options) {
            var blocks = new List<IMarkdownBlock>();
            var elements = cell.Elements;
            AppendBlocksFromElements(elements, block => blocks.Add(block), options, CancellationToken.None, allowQuoteHeuristic: true, trimBoundaryWhitespace: false);

            return new OmdTableCell(blocks);
        }

        private static ColumnAlignment GetAlignment(WordTableCell cell) {
            var alignment = cell.Paragraphs.FirstOrDefault()?.ParagraphAlignment;
            if (alignment == JustificationValues.Center) {
                return ColumnAlignment.Center;
            }

            if (alignment == JustificationValues.Right || alignment == JustificationValues.End) {
                return ColumnAlignment.Right;
            }

            if (alignment == JustificationValues.Left || alignment == JustificationValues.Start) {
                return ColumnAlignment.Left;
            }

            return ColumnAlignment.None;
        }

        private static IMarkdownBlock WrapQuotedBlock(IMarkdownBlock block, int depth) {
            IMarkdownBlock current = block;
            for (int i = 0; i < depth; i++) {
                var quote = new QuoteBlock();
                quote.Children.Add(current);
                current = quote;
            }

            return current;
        }

        private void AppendFootnotes(WordDocument document, MarkdownDoc markdown, WordToMarkdownOptions options) {
            foreach (var footnote in document.FootNotes.OrderBy(fn => fn.ReferenceId)) {
                if (!footnote.ReferenceId.HasValue) {
                    continue;
                }

                var blocks = new List<IMarkdownBlock>();
                foreach (var paragraph in footnote.Paragraphs ?? Enumerable.Empty<WordParagraph>()) {
                    bool hasRuns = false;
                    try {
                        hasRuns = paragraph.GetRuns().Any();
                    } catch (InvalidOperationException ex) {
                        Debug.WriteLine($"GetRuns() failed for footnote paragraph during Markdown AST conversion: {ex.Message}");
                        hasRuns = false;
                    }

                    if (hasRuns && !paragraph.IsFirstRun) {
                        continue;
                    }

                    var paragraphBlocks = BuildParagraphBlocks(paragraph, options, hasCheckbox: false, checkboxChecked: false, allowQuoteHeuristic: false, trimBoundaryWhitespace: false);
                    for (int i = 0; i < paragraphBlocks.Count; i++) {
                        blocks.Add(paragraphBlocks[i]);
                    }
                }

                if (blocks.Count == 0) {
                    continue;
                }

                string label = footnote.ReferenceId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string text = string.Join("\n\n", blocks.Select(block => block.RenderMarkdown()));
                markdown.Add(new FootnoteDefinitionBlock(label, text, blocks, syntaxChildren: null));
            }
        }
    }
}
