using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static bool TryRenderHtmlFallbackViaMarkdownAst(
            string html,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (string.IsNullOrWhiteSpace(html)) {
                return true;
            }

            Omd.MarkdownDoc htmlDocument;
            try {
                var htmlOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
                htmlOptions.PreserveUnsupportedBlocks = false;
                htmlOptions.PreserveUnsupportedInlineHtml = false;
                htmlDocument = html.LoadFromHtml(htmlOptions);
            } catch {
                return false;
            }

            if (htmlDocument.DocumentHeader != null) {
                RenderSharedBlockOmd(
                    htmlDocument.DocumentHeader,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }

            var renderedAny = false;
            foreach (var block in htmlDocument.Blocks) {
                if (block == null) {
                    continue;
                }

                renderedAny = true;
                RenderSharedBlockOmd(
                    block,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }

            return renderedAny;
        }

        private static bool TryRenderWordHeaderFooterSemanticBlock(
            Omd.SemanticFencedBlock block,
            IWordBlockRenderHost currentHost,
            MarkdownToWordOptions options,
            WordDocument document,
            double pageContentWidthPixels) {
            if (block == null || currentHost is not DocumentWordBlockRenderHost) {
                return false;
            }

            if (!TryResolveWordHeaderFooterTarget(block, document, options, out var target)) {
                return false;
            }

            var targetHost = new HeaderFooterWordBlockRenderHost(target);
            if (!string.IsNullOrWhiteSpace(block.Content)) {
                var readerOptions = CreateEffectiveReaderOptions(options);
                readerOptions.FrontMatter = false;
                var fragment = Omd.MarkdownReader.Parse(block.Content, readerOptions);

                if (fragment.DocumentHeader != null) {
                    RenderSharedBlockOmd(
                        fragment.DocumentHeader,
                        targetHost,
                        options,
                        document,
                        currentList: null,
                        listLevel: 0,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: Omd.ColumnAlignment.None);
                }

                foreach (var nested in fragment.Blocks ?? Array.Empty<Omd.IMarkdownBlock>()) {
                    RenderSharedBlockOmd(
                        nested,
                        targetHost,
                        options,
                        document,
                        currentList: null,
                        listLevel: 0,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: Omd.ColumnAlignment.None);
                }
            }

            if (!string.IsNullOrWhiteSpace(block.Caption)) {
                target.AddParagraph(block.Caption!);
            }

            return true;
        }

        private static bool TryRenderWordPageBreakSemanticBlock(
            Omd.SemanticFencedBlock block,
            IWordBlockRenderHost host,
            int quoteDepth,
            Omd.ColumnAlignment alignment) {
            if (block == null ||
                !string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.PageBreakSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            var paragraph = host.CreateParagraph();
            ApplyBlockParagraphFormatting(paragraph, quoteDepth, alignment);
            paragraph.AddBreak(BreakValues.Page);
            if (!string.IsNullOrWhiteSpace(block.Caption)) {
                var captionParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(captionParagraph, quoteDepth, alignment);
                captionParagraph.AddText(block.Caption!);
            }

            return true;
        }

        private static bool TryResolveWordHeaderFooterTarget(
            Omd.SemanticFencedBlock block,
            WordDocument document,
            MarkdownToWordOptions options,
            out WordHeaderFooter target) {
            target = null!;
            if (block == null) {
                return false;
            }

            bool isHeader;
            if (string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.HeaderSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                isHeader = true;
            } else if (string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.FooterSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                isHeader = false;
            } else {
                return false;
            }

            int sectionNumber = 1;
            if (block.FenceInfo.TryGetInt32Attribute("section", out var parsedSection) && parsedSection > 0) {
                sectionNumber = parsedSection;
            }

            if (sectionNumber != 1) {
                options.OnWarning?.Invoke($"Semantic {block.SemanticKind} block requested section {sectionNumber}, but MarkdownToWord currently restores headers and footers only for section 1.");
            }

            var slot = block.FenceInfo.GetAttribute("slot");
            var type = ResolveHeaderFooterType(slot, options, block);
            target = isHeader
                ? document.Sections[0].GetOrCreateHeader(type)
                : document.Sections[0].GetOrCreateFooter(type);
            return true;
        }

        private static HeaderFooterValues ResolveHeaderFooterType(
            string? slot,
            MarkdownToWordOptions options,
            Omd.SemanticFencedBlock block) {
            if (string.IsNullOrWhiteSpace(slot)) {
                return HeaderFooterValues.Default;
            }

            var normalizedSlot = (slot ?? string.Empty).Trim().ToLowerInvariant();
            return normalizedSlot switch {
                "default" => HeaderFooterValues.Default,
                "odd" => HeaderFooterValues.Default,
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                _ => WarnAndReturnDefault(normalizedSlot, options, block)
            };
        }

        private static HeaderFooterValues WarnAndReturnDefault(
            string slot,
            MarkdownToWordOptions options,
            Omd.SemanticFencedBlock block) {
            options.OnWarning?.Invoke($"Semantic {block.SemanticKind} block requested unsupported slot '{slot}'. Falling back to default header or footer.");
            return HeaderFooterValues.Default;
        }

        private static void ProcessTableCellBlocksOmd(
            Omd.TableCell? tableCell,
            WordTableCell wordCell,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (tableCell == null || tableCell.Blocks.Count == 0) {
                return;
            }

            var host = new TableCellWordBlockRenderHost(wordCell);
            RenderSharedBlocksOmd(tableCell.Blocks, host, options, document, quoteDepth: quoteDepth, pageContentWidthPixels: pageContentWidthPixels, alignment: alignment);
        }

        private static void RenderSharedBlocksOmd(
            IEnumerable<Omd.IMarkdownBlock> blocks,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0,
            Omd.ColumnAlignment alignment = Omd.ColumnAlignment.None) {
            if (blocks == null) {
                return;
            }

            foreach (var block in blocks) {
                if (block == null) {
                    continue;
                }

                RenderSharedBlockOmd(block, host, options, document, currentList: null, listLevel, quoteDepth, pageContentWidthPixels, alignment);
            }
        }

        private static void RenderSharedDefinitionListEntryOmd(
            Omd.DefinitionListEntry entry,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            if (entry == null) {
                return;
            }

            bool hasTerm = !string.IsNullOrWhiteSpace(entry.TermMarkdown);
            int nextDefinitionBlockIndex = 0;
            WordParagraph? leadParagraph = null;

            if (hasTerm || (entry.DefinitionBlocks.Count > 0 && entry.DefinitionBlocks[0] is Omd.ParagraphBlock)) {
                leadParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
            }

            if (hasTerm && leadParagraph != null) {
                ProcessInlinesOmd(entry.Term, leadParagraph, options, document, _currentFootnotes);
            }

            if (entry.DefinitionBlocks.Count > 0 && entry.DefinitionBlocks[0] is Omd.ParagraphBlock firstParagraph) {
                if (leadParagraph == null) {
                    leadParagraph = host.CreateParagraph();
                    ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
                }

                if (hasTerm) {
                    var separator = leadParagraph.AddText(": ");
                    var defaultFont = ResolveDefaultFontFamily(options);
                    if (!string.IsNullOrEmpty(defaultFont)) {
                        separator.SetFontFamily(defaultFont!);
                    }
                }

                ProcessInlinesOmd(firstParagraph.Inlines, leadParagraph, options, document, _currentFootnotes);
                nextDefinitionBlockIndex = 1;
            }

            if (entry.DefinitionBlocks.Count == 0 && hasTerm && leadParagraph == null) {
                leadParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(leadParagraph, quoteDepth, alignment);
                ProcessInlinesOmd(entry.Term, leadParagraph, options, document, _currentFootnotes);
            }

            for (int i = nextDefinitionBlockIndex; i < entry.DefinitionBlocks.Count; i++) {
                RenderSharedBlockOmd(
                    entry.DefinitionBlocks[i],
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: quoteDepth,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: alignment);
            }
        }

        private static void RenderSharedTableBlockOmd(
            Omd.TableBlock table,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            double pageContentWidthPixels) {
            var headerCells = table.HeaderCells;
            var rowCells = table.RowCells;
            var cols = headerCells.Count > 0
                ? headerCells.Count
                : (rowCells.Count > 0 ? rowCells[0].Count : 1);
            var rows = rowCells.Count + (headerCells.Count > 0 ? 1 : 0);
            var wordTable = host.CreateTable(rows, cols);
            int rowIndex = 0;

            if (headerCells.Count > 0) {
                for (int columnIndex = 0; columnIndex < headerCells.Count; columnIndex++) {
                    var alignment = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : Omd.ColumnAlignment.None;
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        headerCells[columnIndex].Blocks,
                        cellHost,
                        options,
                        document,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: alignment);
                }
                rowIndex++;
            }

            for (int sourceRowIndex = 0; sourceRowIndex < rowCells.Count; sourceRowIndex++) {
                var row = rowCells[sourceRowIndex];
                for (int columnIndex = 0; columnIndex < row.Count && columnIndex < wordTable.Rows[rowIndex].Cells.Count; columnIndex++) {
                    var alignment = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : Omd.ColumnAlignment.None;
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        row[columnIndex].Blocks,
                        cellHost,
                        options,
                        document,
                        quoteDepth: 0,
                        pageContentWidthPixels: pageContentWidthPixels,
                        alignment: alignment);
                }
                rowIndex++;
            }

            ApplyTableTheme(wordTable, options, headerCells.Count > 0);
        }

        private static void RenderSharedCalloutBlockOmd(
            Omd.CalloutBlock callout,
            IWordBlockRenderHost host,
            MarkdownToWordOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.ColumnAlignment alignment) {
            var titleParagraph = host.CreateParagraph();
            ApplyBlockParagraphFormatting(titleParagraph, quoteDepth, alignment);
            if (callout.TitleInlines != null && (callout.TitleInlines.Items?.Count ?? 0) > 0) {
                ProcessInlinesOmd(callout.TitleInlines, titleParagraph, options, document, _currentFootnotes);
                foreach (var run in titleParagraph.GetRuns()) {
                    run.SetBold();
                }
            } else {
                titleParagraph.AddFormattedText(callout.Title, bold: true);
            }

            ApplyCalloutTitleTheme(titleParagraph, options);

            if (callout.ChildBlocks.Count > 0) {
                RenderSharedBlocksOmd(callout.ChildBlocks, host, options, document, quoteDepth: quoteDepth, pageContentWidthPixels: pageContentWidthPixels, alignment: alignment);
            } else if (!string.IsNullOrWhiteSpace(callout.Body)) {
                var bodyParagraph = host.CreateParagraph();
                ApplyBlockParagraphFormatting(bodyParagraph, quoteDepth, alignment);
                bodyParagraph.AddText(callout.Body);
            }
        }
    }
}
