using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private bool TryRenderHtmlFallbackViaMarkdownAst(
            string html,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.MarkdownColumnAlignment alignment) {
            if (string.IsNullOrWhiteSpace(html)) {
                return true;
            }

            Omd.MarkdownDocument? htmlDocument = options.RawHtmlParser?.Invoke(html, options.DiagnosticSink);
            if (htmlDocument == null) return false;

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

        private bool TryRenderWordHeaderFooterSemanticBlock(
            Omd.SemanticFencedBlock block,
            IWordBlockRenderHost currentHost,
            WordMarkdownImportOptions options,
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
                var fragment = Omd.MarkdownDocument.Parse(block.Content, readerOptions);

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
                        alignment: Omd.MarkdownColumnAlignment.None);
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
                        alignment: Omd.MarkdownColumnAlignment.None);
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
            Omd.MarkdownColumnAlignment alignment) {
            if (block == null ||
                !string.Equals(block.SemanticKind, WordMarkdownSemanticBlocks.PageBreakSemanticKind, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            var paragraph = host.CreateParagraph();
            ApplyBlockParagraphFormatting(paragraph, quoteDepth, alignment);
            paragraph.AddBreak(WordBreakType.Page);
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
            WordMarkdownImportOptions options,
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

        private static WordHeaderFooterType ResolveHeaderFooterType(
            string? slot,
            WordMarkdownImportOptions options,
            Omd.SemanticFencedBlock block) {
            if (string.IsNullOrWhiteSpace(slot)) {
                return WordHeaderFooterType.Default;
            }

            var normalizedSlot = (slot ?? string.Empty).Trim().ToLowerInvariant();
            return normalizedSlot switch {
                "default" => WordHeaderFooterType.Default,
                "odd" => WordHeaderFooterType.Default,
                "first" => WordHeaderFooterType.First,
                "even" => WordHeaderFooterType.Even,
                _ => WarnAndReturnDefault(normalizedSlot, options, block)
            };
        }

        private static WordHeaderFooterType WarnAndReturnDefault(
            string slot,
            WordMarkdownImportOptions options,
            Omd.SemanticFencedBlock block) {
            options.OnWarning?.Invoke($"Semantic {block.SemanticKind} block requested unsupported slot '{slot}'. Falling back to default header or footer.");
            return WordHeaderFooterType.Default;
        }

        private void ProcessTableCellBlocksOmd(
            Omd.TableCell? tableCell,
            WordTableCell wordCell,
            WordMarkdownImportOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.MarkdownColumnAlignment alignment) {
            if (tableCell == null || tableCell.ChildBlocks.Count == 0) {
                return;
            }

            var host = new TableCellWordBlockRenderHost(wordCell);
            RenderSharedBlocksOmd(tableCell.ChildBlocks, host, options, document, quoteDepth: quoteDepth, pageContentWidthPixels: pageContentWidthPixels, alignment: alignment);
        }

        private void RenderSharedBlocksOmd(
            IEnumerable<Omd.IMarkdownBlock> blocks,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
            WordDocument document,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0,
            Omd.MarkdownColumnAlignment alignment = Omd.MarkdownColumnAlignment.None) {
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

        private void RenderSharedDefinitionListEntryOmd(
            Omd.DefinitionListEntry entry,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.MarkdownColumnAlignment alignment) {
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

        private void RenderSharedTableBlockOmd(
            Omd.TableBlock table,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
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
            bool[] numericColumns = ResolveNumericColumns(table, cols);
            Omd.MarkdownColumnAlignment ColumnAlignment(int columnIndex) {
                var authored = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : Omd.MarkdownColumnAlignment.None;
                return authored == Omd.MarkdownColumnAlignment.None && numericColumns[columnIndex] ? Omd.MarkdownColumnAlignment.Right : authored;
            }

            if (headerCells.Count > 0) {
                for (int columnIndex = 0; columnIndex < headerCells.Count; columnIndex++) {
                    var alignment = ColumnAlignment(columnIndex);
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        headerCells[columnIndex].ChildBlocks,
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
                    var alignment = ColumnAlignment(columnIndex);
                    var cellHost = new TableCellWordBlockRenderHost(wordTable.Rows[rowIndex].Cells[columnIndex]);
                    RenderSharedBlocksOmd(
                        row[columnIndex].ChildBlocks,
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
            ApplyContentProportionalColumnWidths(wordTable, pageContentWidthPixels);
        }

        /// <summary>
        /// Sizes columns from their content, the way an HTML auto table layout does: every column gets at least
        /// its longest word, and the remaining width goes to the columns with the most text. Equal-width grids
        /// make short columns waste space while long descriptions wrap into tall rows.
        /// </summary>
        // Numeric columns are right-aligned unless the Markdown table already sets an alignment.
        private static bool[] ResolveNumericColumns(Omd.TableBlock table, int columns) {
            var numeric = new bool[Math.Max(0, columns)];
            for (int column = 0; column < numeric.Length; column++) {
                bool any = false;
                bool all = true;
                foreach (var row in table.Rows) {
                    string value = column < row.Count ? (row[column] ?? string.Empty).Replace("\\", string.Empty).Replace("*", string.Empty) : string.Empty;
                    if (value.Trim().Length == 0) continue;
                    any = true;
                    if (!IsNumericCellText(value)) { all = false; break; }
                }
                numeric[column] = any && all;
            }
            return numeric;
        }

        private static bool IsNumericCellText(string value) {
            string text = value.Trim();
            if (text.Length == 0 || text.Length > 40) return false;
            int start = 0, end = text.Length;
            if (text[start] == '+' || text[start] == '-' || text[start] == '\u2212' || text[start] == '(') start++;
            while (start < end && "$\u20AC\u00A3\u00A5".IndexOf(text[start]) >= 0) start++;
            while (end > start && (text[end - 1] == '%' || text[end - 1] == ')')) end--;
            bool digit = false, dot = false;
            for (int index = start; index < end; index++) {
                char c = text[index];
                if (char.IsDigit(c)) { digit = true; continue; }
                if (c == '.' && !dot) { dot = true; continue; }
                if ((c == ',' || c == ' ') && index > start && index + 1 < end && char.IsDigit(text[index + 1])) continue;
                return false;
            }
            return digit && start < end;
        }

        private static bool IsMonospaceFont(string? font) {
            if (string.IsNullOrWhiteSpace(font)) return false;
            return string.Equals(font, WordFontResolver.Resolve("monospace"), StringComparison.OrdinalIgnoreCase)
                || font!.IndexOf("Mono", StringComparison.OrdinalIgnoreCase) >= 0
                || font.IndexOf("Consol", StringComparison.OrdinalIgnoreCase) >= 0
                || font.IndexOf("Courier", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(font, "Menlo", StringComparison.OrdinalIgnoreCase)
                || string.Equals(font, "Monaco", StringComparison.OrdinalIgnoreCase);
        }

        private static void ApplyContentProportionalColumnWidths(WordTable table, double pageContentWidthPixels) {
            if (table.Rows.Count == 0) {
                return;
            }

            int columns = table.Rows[0].Cells.Count;
            if (columns < 2 || table.Rows.Any(row => row.Cells.Count != columns)) {
                return;
            }

            var minimum = new double[columns];
            var maximum = new double[columns];
            // Short values (a setting name, a default) read badly when wrapped, so they count as one word.
            const double ShortValueCharacters = 16D;
            // Monospace glyphs are wider than proportional body text.
            const double MonospaceWidthFactor = 1.3D;
            foreach (var row in table.Rows) {
                for (int column = 0; column < columns; column++) {
                    double longestWord = 0;
                    double cellTotal = 0;
                    foreach (var paragraph in row.Cells[column]._tableCell.Elements<Paragraph>()) {
                        double current = 0;
                        foreach (var run in paragraph.Descendants<Run>()) {
                            string runText = string.Concat(run.Elements<Text>().Select(text => text.Text));
                            double factor = IsMonospaceFont(run.RunProperties?.RunFonts?.Ascii?.Value) ? MonospaceWidthFactor : 1D;
                            foreach (char character in runText) {
                                if (char.IsWhiteSpace(character)) {
                                    longestWord = Math.Max(longestWord, current);
                                    current = 0;
                                } else {
                                    current += factor;
                                }
                            }
                            cellTotal += runText.Length * factor;
                        }
                        longestWord = Math.Max(longestWord, current);
                        cellTotal += 1D;
                    }

                    if (cellTotal <= ShortValueCharacters) longestWord = Math.Max(longestWord, cellTotal);
                    minimum[column] = Math.Max(minimum[column], longestWord);
                    maximum[column] = Math.Max(maximum[column], cellTotal);
                }
            }

            // Approximate character units: about 7 px per character at body size, plus two characters of cell padding.
            const double CellPaddingCharacters = 2D;
            double available = Math.Max(columns * 4D, pageContentWidthPixels / 7D);
            for (int column = 0; column < columns; column++) {
                minimum[column] = Math.Max(3D, minimum[column]) + CellPaddingCharacters;
                maximum[column] = Math.Max(minimum[column], maximum[column] + CellPaddingCharacters);
            }

            double minimumTotal = minimum.Sum();
            double maximumTotal = maximum.Sum();
            var widths = new double[columns];
            if (maximumTotal <= available) {
                for (int column = 0; column < columns; column++) widths[column] = maximum[column];
            } else if (minimumTotal >= available) {
                for (int column = 0; column < columns; column++) widths[column] = minimum[column];
            } else {
                double slack = available - minimumTotal;
                double flexible = maximumTotal - minimumTotal;
                for (int column = 0; column < columns; column++) {
                    widths[column] = minimum[column] + slack * (maximum[column] - minimum[column]) / flexible;
                }
            }

            double total = widths.Sum();
            if (total <= 0) {
                return;
            }

            int[] percentages = widths.Select(width => Math.Max(1, (int)Math.Round(width * 1000D / total))).ToArray();
            table.SetColumnWidthsPercentage(percentages);
            // Keep the grid in step with the cell widths: fixed-layout consumers such as LibreOffice read tblGrid.
            double contentTwips = Math.Max(1D, pageContentWidthPixels) * TwipsPerInch / PixelsPerInch;
            int percentTotal = percentages.Sum();
            table.GridColumnWidth = percentages.Select(value => (int)Math.Round(contentTwips * value / percentTotal)).ToList();
        }

        private void RenderSharedCalloutBlockOmd(
            Omd.CalloutBlock callout,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
            WordDocument document,
            int quoteDepth,
            double pageContentWidthPixels,
            Omd.MarkdownColumnAlignment alignment) {
            var titleParagraph = host.CreateParagraph();
            ApplyBlockParagraphFormatting(titleParagraph, quoteDepth, alignment);
            if (callout.TitleInlines != null && callout.TitleInlines.Nodes.Count > 0) {
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
                ApplyBodyTextTheme(bodyParagraph, options);
            }
        }
    }
}
