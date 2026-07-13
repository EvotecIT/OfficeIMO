using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderTable(PdfCore.PdfDocument pdf, TableBlock table, MarkdownPdfStyle visualTheme) {
        int columnCount = GetMarkdownTableColumnCount(table);
        if (columnCount == 0) {
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        bool hasSpans = TableHasSpans(table);
        bool useStructuredCells = hasSpans || TableHasCellTextStyles(table);
        if (table.Headers.Count > 0) {
            rows.Add(useStructuredCells
                ? CreateTableRow(table.HeaderCells, columnCount, visualTheme, preserveSpans: hasSpans)
                : CreateTableRow(table.HeaderInlines, columnCount, visualTheme));
        }

        if (useStructuredCells) {
            IReadOnlyList<IReadOnlyList<TableCell>> bodyCells = table.RowCells;
            for (int rowIndex = 0; rowIndex < bodyCells.Count; rowIndex++) {
                rows.Add(CreateTableRow(bodyCells[rowIndex], columnCount, visualTheme, preserveSpans: hasSpans));
            }
        } else {
            IReadOnlyList<IReadOnlyList<InlineSequence>> bodyRows = table.RowInlines;
            for (int rowIndex = 0; rowIndex < bodyRows.Count; rowIndex++) {
                rows.Add(CreateTableRow(bodyRows[rowIndex], columnCount, visualTheme));
            }
        }

        PdfCore.PdfTableStyle style = visualTheme.TableStyleSnapshot;
        style.HeaderRowCount = table.Headers.Count > 0 ? 1 : 0;
        style.RepeatHeaderRowCount = table.Headers.Count > 0 ? 1 : 0;
        style.RightAlignNumeric = true;
        ApplyMarkdownTableAlignments(style, table, columnCount, rows.Count);
        ApplyMarkdownTableCellAlignments(style, table);
        ApplyMarkdownTableColumnWidths(style, table, columnCount);
        ApplyMarkdownTableCellBackgrounds(style, table, columnCount);

        pdf.Table(rows, style: style);
    }

    private static int GetMarkdownTableColumnCount(TableBlock table) {
        int rowColumnCount = table.Rows.Count == 0 ? 0 : table.Rows.Max(row => Math.Min(row.Count, TableBlock.MaxEffectiveColumnCount));
        int columnCount = Math.Max(Math.Min(table.Headers.Count, TableBlock.MaxEffectiveColumnCount), rowColumnCount);
        columnCount = Math.Max(columnCount, Math.Min(table.Alignments.Count, TableBlock.MaxEffectiveColumnCount));
        columnCount = Math.Max(columnCount, Math.Min(table.ColumnWidthPoints.Count, TableBlock.MaxEffectiveColumnCount));
        columnCount = Math.Max(columnCount, Math.Min(table.ColumnWidthWeights.Count, TableBlock.MaxEffectiveColumnCount));
        columnCount = Math.Max(columnCount, CountLogicalColumns(table.HeaderCells));

        IReadOnlyList<IReadOnlyList<TableCell>> rowCells = table.RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            columnCount = Math.Max(columnCount, CountLogicalColumns(rowCells[rowIndex]));
        }

        return Math.Min(columnCount, TableBlock.MaxEffectiveColumnCount);
    }

    private static int CountLogicalColumns(IReadOnlyList<TableCell> cells) {
        int count = 0;
        int meaningfulCount = GetMeaningfulCellCount(cells);
        for (int cellIndex = 0; cellIndex < meaningfulCount; cellIndex++) {
            count += Math.Max(1, cells[cellIndex]?.ColumnSpan ?? 1);
            if (count >= TableBlock.MaxEffectiveColumnCount) {
                return TableBlock.MaxEffectiveColumnCount;
            }
        }

        return count;
    }

    private static bool TableHasSpans(TableBlock table) {
        if (RowHasSpans(table.HeaderCells)) {
            return true;
        }

        IReadOnlyList<IReadOnlyList<TableCell>> rowCells = table.RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            if (RowHasSpans(rowCells[rowIndex])) {
                return true;
            }
        }

        return false;
    }

    private static bool TableHasCellTextStyles(TableBlock table) {
        if (RowHasCellTextStyles(table.HeaderCells)) {
            return true;
        }

        IReadOnlyList<IReadOnlyList<TableCell>> rowCells = table.RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            if (RowHasCellTextStyles(rowCells[rowIndex])) {
                return true;
            }
        }

        return false;
    }

    private static bool RowHasCellTextStyles(IReadOnlyList<TableCell> cells) {
        int meaningfulCount = GetMeaningfulCellCount(cells);
        for (int cellIndex = 0; cellIndex < meaningfulCount; cellIndex++) {
            if (CellHasTextStyle(cells[cellIndex])) {
                return true;
            }
        }

        return false;
    }

    private static bool CellHasTextStyle(TableCell? cell) {
        return cell != null &&
            (!string.IsNullOrWhiteSpace(cell.TextColor) ||
             cell.Bold ||
             cell.Italic ||
             cell.Underline ||
             cell.Strikethrough);
    }

    private static bool RowHasSpans(IReadOnlyList<TableCell> cells) {
        int meaningfulCount = GetMeaningfulCellCount(cells);
        for (int cellIndex = 0; cellIndex < meaningfulCount; cellIndex++) {
            TableCell? cell = cells[cellIndex];
            if (cell != null && (cell.ColumnSpan > 1 || cell.RowSpan > 1)) {
                return true;
            }
        }

        return false;
    }

    private static int GetMeaningfulCellCount(IReadOnlyList<TableCell> cells) {
        int count = cells?.Count ?? 0;
        while (count > 0 && IsGeneratedPaddingCell(cells![count - 1])) {
            count--;
        }

        return count;
    }

    private static bool IsGeneratedPaddingCell(TableCell? cell) {
        return cell == null || (cell.ChildBlocks.Count == 0 && cell.ColumnSpan == 1 && cell.RowSpan == 1);
    }

    private static void ApplyMarkdownTableColumnWidths(PdfCore.PdfTableStyle style, TableBlock table, int columnCount) {
        if (columnCount == 0) {
            return;
        }

        bool hasFixedWidths = false;
        if (table.ColumnWidthPoints.Count > 0) {
            var fixedWidths = new List<double?>(columnCount);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                double? width = columnIndex < table.ColumnWidthPoints.Count ? table.ColumnWidthPoints[columnIndex] : null;
                fixedWidths.Add(width.HasValue && width.Value > 0D && !double.IsNaN(width.Value) && !double.IsInfinity(width.Value) ? width : null);
                hasFixedWidths |= fixedWidths[columnIndex].HasValue;
            }

            if (hasFixedWidths) {
                style.ColumnWidthPoints = fixedWidths;
            }
        }

        if (table.ColumnWidthWeights.Count == 0) {
            return;
        }

        var weights = new List<double>(columnCount);
        bool hasWeights = false;
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            double weight = columnIndex < table.ColumnWidthWeights.Count ? table.ColumnWidthWeights[columnIndex] : 1D;
            if (weight <= 0D || double.IsNaN(weight) || double.IsInfinity(weight)) {
                weight = 1D;
            }

            weights.Add(weight);
            if (Math.Abs(weight - 1D) > 0.001) {
                hasWeights = true;
            }
        }

        if (hasWeights || hasFixedWidths) {
            style.ColumnWidthWeights = weights;
        }
    }

    private static void ApplyMarkdownTableCellBackgrounds(PdfCore.PdfTableStyle style, TableBlock table, int columnCount) {
        if (columnCount <= 0) {
            return;
        }

        var cellFills = style.CellFills == null
            ? new Dictionary<(int Row, int Column), PdfCore.PdfColor>()
            : new Dictionary<(int Row, int Column), PdfCore.PdfColor>(style.CellFills);

        bool hasCellFills = cellFills.Count > 0;
        var activeRowSpans = new int[columnCount];
        if (table.Headers.Count > 0) {
            ApplyMarkdownTableCellBackgrounds(cellFills, table.HeaderCells, pdfRowIndex: 0, columnCount, activeRowSpans, ref hasCellFills);
        }

        int bodyRowOffset = table.Headers.Count > 0 ? 1 : 0;
        IReadOnlyList<IReadOnlyList<TableCell>> rowCells = table.RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            ApplyMarkdownTableCellBackgrounds(cellFills, rowCells[rowIndex], bodyRowOffset + rowIndex, columnCount, activeRowSpans, ref hasCellFills);
        }

        if (hasCellFills) {
            style.CellFills = cellFills;
        }
    }

    private static void ApplyMarkdownTableCellBackgrounds(Dictionary<(int Row, int Column), PdfCore.PdfColor> cellFills, IReadOnlyList<TableCell> cells, int pdfRowIndex, int columnCount, int[] activeRowSpans, ref bool hasCellFills) {
        int meaningfulCount = GetMeaningfulCellCount(cells);
        int logicalColumn = 0;
        for (int cellIndex = 0; cellIndex < meaningfulCount; cellIndex++) {
            while (logicalColumn < columnCount && activeRowSpans[logicalColumn] > 0) {
                logicalColumn++;
            }

            if (logicalColumn >= columnCount) {
                break;
            }

            TableCell? cell = cells[cellIndex];
            if (cell == null) {
                logicalColumn++;
                continue;
            }

            if (TryParseTableCellColor(cell.BackgroundColor, out PdfCore.PdfColor fill)) {
                cellFills[(pdfRowIndex, logicalColumn)] = fill;
                hasCellFills = true;
            }

            int columnSpan = Math.Max(1, Math.Min(cell.ColumnSpan, columnCount - logicalColumn));
            int rowSpan = Math.Max(1, cell.RowSpan);
            for (int column = logicalColumn; column < logicalColumn + columnSpan; column++) {
                activeRowSpans[column] = Math.Max(activeRowSpans[column], rowSpan);
            }

            logicalColumn += columnSpan;
        }

        for (int column = 0; column < activeRowSpans.Length; column++) {
            if (activeRowSpans[column] > 0) {
                activeRowSpans[column]--;
            }
        }
    }

    private static bool TryParseTableCellColor(string? value, out PdfCore.PdfColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value) || !OfficeColor.TryParse(value, out OfficeColor officeColor)) {
            return false;
        }

        PdfCore.PdfColor? pdfColor = PdfCore.PdfColor.FromOfficeColorOrNull(officeColor);
        if (!pdfColor.HasValue) {
            return false;
        }

        color = pdfColor.Value;
        return true;
    }

    private static InlineStyle CreateTableCellInlineStyle(TableCell cell, MarkdownPdfStyle visualTheme) {
        InlineStyle style = CreateInlineStyle(visualTheme).With(
            bold: cell.Bold ? true : null,
            italic: cell.Italic ? true : null,
            underline: cell.Underline ? true : null,
            strike: cell.Strikethrough ? true : null);

        if (TryParseTableCellColor(cell.TextColor, out PdfCore.PdfColor textColor)) {
            style = style.With(color: textColor);
        }

        return style;
    }

    private static void ApplyMarkdownTableAlignments(PdfCore.PdfTableStyle style, TableBlock table, int columnCount, int rowCount) {
        if (table.Alignments.Count == 0 || columnCount == 0 || rowCount == 0) {
            return;
        }

        var alignments = new List<PdfCore.PdfColumnAlign>(columnCount);
        var cellAlignments = style.CellAlignments == null
            ? new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>()
            : new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>(style.CellAlignments);
        bool hasExplicitAlignment = false;

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            ColumnAlignment markdownAlignment = columnIndex < table.Alignments.Count ? table.Alignments[columnIndex] : ColumnAlignment.None;
            PdfCore.PdfColumnAlign pdfAlignment = MapMarkdownTableAlignment(markdownAlignment);
            alignments.Add(pdfAlignment);

            if (markdownAlignment == ColumnAlignment.None) {
                continue;
            }

            hasExplicitAlignment = true;
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                cellAlignments[(rowIndex, columnIndex)] = pdfAlignment;
            }
        }

        if (!hasExplicitAlignment) {
            return;
        }

        style.Alignments = alignments;
        style.CellAlignments = cellAlignments;
    }

    private static void ApplyMarkdownTableCellAlignments(PdfCore.PdfTableStyle style, TableBlock table) {
        var cellAlignments = style.CellAlignments == null
            ? new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>()
            : new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>(style.CellAlignments);

        bool hasExplicitAlignment = cellAlignments.Count > 0;
        int columnCount = GetMarkdownTableColumnCount(table);
        var activeRowSpans = new int[Math.Max(0, columnCount)];
        ApplyMarkdownTableCellAlignments(cellAlignments, table.HeaderCells, pdfRowIndex: 0, columnCount, activeRowSpans, ref hasExplicitAlignment);

        int bodyRowOffset = table.Headers.Count > 0 ? 1 : 0;
        IReadOnlyList<IReadOnlyList<TableCell>> rowCells = table.RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            ApplyMarkdownTableCellAlignments(cellAlignments, rowCells[rowIndex], bodyRowOffset + rowIndex, columnCount, activeRowSpans, ref hasExplicitAlignment);
        }

        if (hasExplicitAlignment) {
            style.CellAlignments = cellAlignments;
        }
    }

    private static void ApplyMarkdownTableCellAlignments(Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign> cellAlignments, IReadOnlyList<TableCell> cells, int pdfRowIndex, int columnCount, int[] activeRowSpans, ref bool hasExplicitAlignment) {
        int meaningfulCount = GetMeaningfulCellCount(cells);
        int logicalColumn = 0;
        for (int cellIndex = 0; cellIndex < meaningfulCount; cellIndex++) {
            while (logicalColumn < columnCount && activeRowSpans[logicalColumn] > 0) {
                logicalColumn++;
            }

            if (logicalColumn >= columnCount) {
                break;
            }

            TableCell? cell = cells[cellIndex];
            if (cell == null) {
                logicalColumn++;
                continue;
            }

            if (cell.Alignment != ColumnAlignment.None) {
                cellAlignments[(pdfRowIndex, logicalColumn)] = MapMarkdownTableAlignment(cell.Alignment);
                hasExplicitAlignment = true;
            }

            int columnSpan = Math.Max(1, Math.Min(cell.ColumnSpan, columnCount - logicalColumn));
            int rowSpan = Math.Max(1, cell.RowSpan);
            for (int column = logicalColumn; column < logicalColumn + columnSpan; column++) {
                activeRowSpans[column] = Math.Max(activeRowSpans[column], rowSpan);
            }

            logicalColumn += columnSpan;
        }

        for (int column = 0; column < activeRowSpans.Length; column++) {
            if (activeRowSpans[column] > 0) {
                activeRowSpans[column]--;
            }
        }
    }

    private static PdfCore.PdfColumnAlign MapMarkdownTableAlignment(ColumnAlignment alignment) {
        return alignment switch {
            ColumnAlignment.Center => PdfCore.PdfColumnAlign.Center,
            ColumnAlignment.Right => PdfCore.PdfColumnAlign.Right,
            _ => PdfCore.PdfColumnAlign.Left
        };
    }

    private static PdfCore.PdfTableCell[] CreateTableRow(IReadOnlyList<InlineSequence> cells, int columnCount, MarkdownPdfStyle visualTheme) {
        var row = new PdfCore.PdfTableCell[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            InlineSequence? sequence = columnIndex < cells.Count ? cells[columnIndex] : null;
            IReadOnlyList<PdfTextRun> runs = sequence == null ? Array.Empty<PdfTextRun>() : ToTextRuns(sequence, CreateInlineStyle(visualTheme));
            row[columnIndex] = runs.Count == 0
                ? PdfCore.PdfTableCell.TextCell(string.Empty)
                : PdfCore.PdfTableCell.RichTextCell(runs);
        }

        return row;
    }

    private static PdfCore.PdfTableCell[] CreateTableRow(IReadOnlyList<TableCell> cells, int columnCount, MarkdownPdfStyle visualTheme, bool preserveSpans) {
        var row = new List<PdfCore.PdfTableCell>();
        int meaningfulCount = preserveSpans ? GetMeaningfulCellCount(cells) : cells.Count;
        int logicalColumn = 0;
        for (int cellIndex = 0; cellIndex < meaningfulCount && logicalColumn < columnCount; cellIndex++) {
            TableCell cell = cells[cellIndex] ?? new TableCell();
            IReadOnlyList<PdfTextRun> runs = CreateTableCellRuns(cell, visualTheme);
            int columnSpan = preserveSpans ? Math.Max(1, Math.Min(cell.ColumnSpan, columnCount - logicalColumn)) : 1;
            int rowSpan = preserveSpans ? Math.Max(1, cell.RowSpan) : 1;
            row.Add(runs.Count == 0
                ? PdfCore.PdfTableCell.Merge(string.Empty, columnSpan, rowSpan)
                : PdfCore.PdfTableCell.Merge(runs, columnSpan, rowSpan));
            logicalColumn += columnSpan;
        }

        if (!preserveSpans) {
            while (row.Count < columnCount) {
                row.Add(PdfCore.PdfTableCell.TextCell(string.Empty));
            }
        }

        return row.ToArray();
    }

    private static IReadOnlyList<PdfTextRun> CreateTableCellRuns(TableCell cell, MarkdownPdfStyle visualTheme) {
        InlineStyle style = CreateTableCellInlineStyle(cell, visualTheme);
        if (cell.ChildBlocks.Count == 1 && cell.ChildBlocks[0] is ParagraphBlock paragraph) {
            return ToTextRuns(paragraph.Inlines, style);
        }

        string markdown = cell.Markdown;
        return string.IsNullOrEmpty(markdown)
            ? Array.Empty<PdfTextRun>()
            : new[] { CreateRun(markdown, style) };
    }

    private static void RenderCodeBlock(PdfCore.PdfDocument pdf, CodeBlock code, MarkdownPdfStyle visualTheme) {
        string content = code.Content ?? string.Empty;
        string language = code.Language?.Trim() ?? string.Empty;

        pdf.PanelParagraph(builder => {
            if (language.Length > 0) {
                builder.FontSize(visualTheme.CodeBlockLabelFontSizeSnapshot)
                    .Bold(language, visualTheme.CodeBlockLabelColorSnapshot)
                    .LineBreak();
            }

            builder.Font(PdfCore.PdfStandardFont.Courier)
                .FontSize(visualTheme.CodeBlockFontSizeSnapshot)
                .Color(visualTheme.CodeBlockTextColorSnapshot);
            AppendTextWithLineBreaks(builder, content);
        }, visualTheme.CodeBlockPanelStyleSnapshot);

        if (!string.IsNullOrWhiteSpace(code.Caption)) {
            pdf.Paragraph(builder => builder.Italic(code.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderSemanticFencedBlock(PdfCore.PdfDocument pdf, SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (TryRenderChartFencedBlock(pdf, semantic, options, visualTheme)) {
            return;
        }

        if (!IsChartSemanticFence(semantic)) {
            AddWarning(options, "UnsupportedSemanticFence", semantic.SemanticKind, "The Markdown semantic fenced block is rendered as a styled code panel because no PDF visual renderer is registered for its semantic kind.");
        }

        string label = semantic.SemanticKind;
        if (!string.IsNullOrWhiteSpace(semantic.Language)) {
            label += " / " + semantic.Language;
        }

        pdf.PanelParagraph(builder => {
            builder.FontSize(visualTheme.SemanticBlockLabelFontSizeSnapshot)
                .Bold(label, visualTheme.SemanticBlockLabelColorSnapshot);
            builder.LineBreak();
            builder.Font(PdfCore.PdfStandardFont.Courier)
                .FontSize(visualTheme.SemanticBlockFontSizeSnapshot)
                .Color(visualTheme.SemanticBlockTextColorSnapshot);
            AppendTextWithLineBreaks(builder, semantic.Content ?? string.Empty);
        }, visualTheme.SemanticBlockPanelStyleSnapshot);

        if (!string.IsNullOrWhiteSpace(semantic.Caption)) {
            pdf.Paragraph(builder => builder.Italic(semantic.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderCalloutBlock(PdfCore.PdfDocument pdf, CalloutBlock callout, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        string title = string.IsNullOrWhiteSpace(callout.Title) ? FormatTitleFromKind(callout.Kind) : callout.Title;

        IReadOnlyList<IMarkdownBlock> children = callout.ChildBlocks;
        bool canRenderChildrenInsidePanel = children.Count > 0 && CanRenderBlocksInsidePanel(children);
        PdfCore.PanelStyle panelStyle = visualTheme.CreateCalloutPanelStyle(callout.Kind);
        if (canRenderChildrenInsidePanel) {
            pdf.Panel(panel => {
                panel.Paragraph(builder => {
                    if (IsEmpty(callout.TitleInlines)) {
                        builder.Bold(title);
                    } else {
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
                    }
                });
                RenderBlocks(pdf, children, document, options, visualTheme);
            }, panelStyle);
            return;
        }

        if (children.Count == 0 && !string.IsNullOrWhiteSpace(callout.Body)) {
            pdf.Panel(panel => {
                panel.Paragraph(builder => {
                    if (IsEmpty(callout.TitleInlines)) {
                        builder.Bold(title);
                    } else {
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
                    }

                    builder.LineBreak();
                    AppendTextWithLineBreaks(builder, callout.Body);
                });
            }, panelStyle);
            return;
        }

        bool renderedInsidePanel = false;
        pdf.PanelParagraph(builder => {
            if (IsEmpty(callout.TitleInlines)) {
                builder.Bold(title);
            } else {
                AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
            }

            if (canRenderChildrenInsidePanel) {
                renderedInsidePanel = TryAppendBlocksInsidePanel(builder, children, CreateInlineStyle(visualTheme), visualTheme, lineBreakBeforeFirst: true);
            } else if (children.Count == 0 && !string.IsNullOrWhiteSpace(callout.Body)) {
                builder.LineBreak();
                AppendTextWithLineBreaks(builder, callout.Body);
                renderedInsidePanel = true;
            }
        }, panelStyle);

        if (children.Count > 0 && !renderedInsidePanel) {
            RenderBlocksWithPanelRuns(
                pdf,
                children,
                document,
                options,
                visualTheme,
                panelStyle,
                panel => panel.Paragraph(builder => {
                    if (IsEmpty(callout.TitleInlines)) {
                        builder.Bold(title);
                    } else {
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
                    }
                }));
        }
    }

    private static void RenderDetailsBlock(PdfCore.PdfDocument pdf, DetailsBlock details, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (details.Summary != null) {
            pdf.PanelParagraph(builder => {
                builder.Bold(details.Open ? "Details: " : "Collapsed details: ");
                AppendInlines(builder, details.Summary.Inlines, CreateInlineStyle(visualTheme).With(bold: true));
            }, visualTheme.DetailsPanelStyleSnapshot);
        }

        RenderBlocks(pdf, details.ChildBlocks, document, options, visualTheme);
    }

    private static void RenderDefinitionList(PdfCore.PdfDocument pdf, DefinitionListBlock definitionList, MarkdownPdfStyle visualTheme) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        if (items.Count == 0) {
            return;
        }

        var rows = new List<PdfCore.PdfKeyValueRow>();
        for (int i = 0; i < items.Count; i++) {
            IReadOnlyList<PdfTextRun> termRuns = ToTextRuns(items[i].Term, CreateInlineStyle(visualTheme).With(bold: true));
            IReadOnlyList<PdfTextRun> definitionRuns = ToTextRuns(items[i].Definition, CreateInlineStyle(visualTheme));
            rows.Add(PdfCore.PdfKeyValueRow.Rich(termRuns, definitionRuns));
        }

        PdfCore.PdfTableStyle style = visualTheme.DefinitionListTableStyleSnapshot;
        pdf.KeyValueTable(rows, style: style);
    }

    private static void RenderFootnoteDefinition(PdfCore.PdfDocument pdf, FootnoteDefinitionBlock footnote, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (string.IsNullOrWhiteSpace(footnote.Label)) {
            return;
        }

        pdf.Paragraph(builder => {
            builder.Superscript(footnote.Label);
            builder.Text(" ");
            if (footnote.ChildBlocks.Count == 1 && footnote.ChildBlocks[0] is ParagraphBlock paragraph) {
                AppendInlines(builder, paragraph.Inlines, CreateInlineStyle(visualTheme));
            } else {
                AppendTextWithLineBreaks(builder, footnote.Text);
            }
        }, style: new PdfCore.PdfParagraphStyle { SpacingBefore = 4, SpacingAfter = 4 });

        if (footnote.ChildBlocks.Count > 1) {
            for (int i = 1; i < footnote.ChildBlocks.Count; i++) {
                RenderBlock(pdf, footnote.ChildBlocks[i], document, options, visualTheme);
            }
        }
    }

    private static void RenderTocBlock(PdfCore.PdfDocument pdf, TocBlock toc, MarkdownPdfStyle visualTheme) {
        if (toc.Entries.Count == 0) {
            return;
        }

        int baseLevel = toc.NormalizeLevels ? toc.Entries.Min(entry => entry.Level) : 1;
        if (ShouldRenderTocAsPanel(toc)) {
            pdf.PanelParagraph(builder => {
                if (toc.IncludeTitle && !string.IsNullOrWhiteSpace(toc.Title)) {
                    builder.Bold(true)
                        .FontSize(GetTocTitleFontSize(toc.TitleLevel))
                        .Color(visualTheme.TocTitleColorSnapshot)
                        .Text(toc.Title.Trim())
                        .Bold(false)
                        .ResetFontSize()
                        .Color(visualTheme.TocTextColorSnapshot);

                    if (toc.Entries.Count > 0) {
                        builder.LineBreak();
                        builder.LineBreak();
                    }
                } else {
                    builder.Color(visualTheme.TocTextColorSnapshot);
                }

                for (int i = 0; i < toc.Entries.Count; i++) {
                    if (i > 0) {
                        builder.LineBreak();
                    }

                    AppendTocEntry(builder, toc.Entries[i], i, baseLevel, toc.Ordered, visualTheme);
                }
            }, visualTheme.TocPanelStyleSnapshot, defaultColor: visualTheme.TocTextColorSnapshot);
            return;
        }

        var items = new List<PdfCore.PdfListItem>();
        for (int i = 0; i < toc.Entries.Count; i++) {
            TocBlock.Entry entry = toc.Entries[i];
            string prefix = new string(' ', Math.Max(0, entry.Level - baseLevel) * 2);
            string text = prefix + entry.Text;
            items.Add(string.IsNullOrWhiteSpace(entry.Anchor)
                ? PdfCore.PdfListItem.Plain(text)
                : PdfCore.PdfListItem.Rich(new[] { PdfTextRun.LinkToBookmark(text, entry.Anchor, visualTheme.TocLinkColorSnapshot, underline: true, contents: "Table of contents: " + entry.Text) }));
        }

        if (toc.Ordered) {
            pdf.RichNumbered(items);
        } else {
            pdf.RichBullets(items);
        }
    }

    private static bool ShouldRenderTocAsPanel(TocBlock toc) =>
        toc.Layout != TocLayout.List || toc.Chrome == TocChrome.Panel || toc.Chrome == TocChrome.Outline;

    private static void AppendTocEntry(PdfCore.PdfParagraphBuilder builder, TocBlock.Entry entry, int index, int baseLevel, bool ordered, MarkdownPdfStyle visualTheme) {
        int depth = Math.Max(0, entry.Level - baseLevel);
        if (depth > 0) {
            builder.Text(new string(' ', depth * 4));
        }

        string marker = ordered ? (index + 1).ToString(CultureInfo.InvariantCulture) + ". " : "• ";
        builder.Color(visualTheme.TocTextColorSnapshot).Text(marker);
        string text = string.IsNullOrWhiteSpace(entry.Text) ? "(untitled)" : entry.Text.Trim();
        if (string.IsNullOrWhiteSpace(entry.Anchor)) {
            builder.Color(visualTheme.TocLinkColorSnapshot).Text(text);
        } else {
            builder.LinkToBookmark(text, entry.Anchor, visualTheme.TocLinkColorSnapshot, underline: false, contents: "Table of contents: " + text);
        }

        builder.Color(visualTheme.TocTextColorSnapshot);
    }

    private static double GetTocTitleFontSize(int titleLevel) {
        return titleLevel <= 1 ? 15D : titleLevel == 2 ? 13D : 11.5D;
    }
}
