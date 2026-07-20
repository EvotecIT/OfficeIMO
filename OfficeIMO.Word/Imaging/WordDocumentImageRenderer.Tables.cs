using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool AddTable(
            WordTable table,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers = null,
            bool allowNestedTable = false) {
            context.ThrowIfCancellationRequested();
            List<WordTableRow> rows = table.Rows;
            if (rows.Count == 0) {
                return false;
            }

            if (table.IsNestedTable && !allowNestedTable) {
                AddDiagnostic(diagnostics, "unsupported-word-nested-table", "Skipped a nested Word table because dependency-free nested table layout is not implemented yet.");
                return false;
            }

            int columnCount = Math.Max(1, rows.Max(row => row.Cells.Count));
            double[] columnWidths = ResolveColumnWidths(table, columnCount, context.ContentWidth);
            double[] rowHeights = ResolveRowHeights(
                rows,
                columnWidths,
                listMarkers,
                context.CancellationToken,
                context.CancellationCheckpoint);
            double tableHeight = rowHeights.Sum();
            double remainingHeight = Math.Max(0D, context.ContentBottom - context.Y);
            if ((tableHeight > remainingHeight || HasTableRowStartBreak(table, rows)) && context.CanAdvancePageForOverflow) {
                return AddPaginatedTableRows(table, rows, columnWidths, rowHeights, GetRepeatingHeaderRowCount(rows), context, diagnostics, listMarkers);
            }

            if (!EnsureVerticalSpace(context, tableHeight, diagnostics)) {
                return false;
            }

            double tableWidth = columnWidths.Sum();
            double tableLeft = ResolveTableLeft(table, context.Left, context.ContentWidth, tableWidth);
            double rowTop = context.Y;
            if (context.IsTargetPage) {
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    context.ThrowIfCancellationRequested();
                    AddTableRow(context, table, rows, rowIndex, tableLeft, rowTop, columnWidths, rowHeights, diagnostics, listMarkers);
                    rowTop += rowHeights[rowIndex];
                }
            }

            context.Y += tableHeight + ParagraphGapPoints;
            return true;
        }

        private static bool AddPaginatedTableRows(
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            int repeatingHeaderRowCount,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            double tableWidth = columnWidths.Sum();
            bool consumedRows = false;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                context.ThrowIfCancellationRequested();
                if (!TryAdvanceForTableRowStartPageBreak(table, rows, rowIndex, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                    return consumedRows;
                }

                double rowHeight = rowHeights[rowIndex];
                if (rowHeight > context.ContentHeight) {
                    if (TryAddPaginatedSplitTableRow(table, rows, rowIndex, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                        consumedRows = true;
                        continue;
                    }

                    if (context.IsTargetPage) {
                        AddDiagnostic(diagnostics, "unsupported-word-table-row-pagination", "Skipped a Word table row because splitting an individual row across pages is not implemented yet.");
                    }

                    return consumedRows;
                }

                if (context.Y + rowHeight > context.ContentBottom) {
                    if (context.Y > context.Top && context.CanAdvancePageForOverflow) {
                        if (rows[rowIndex].AllowRowToBreakAcrossPages &&
                            TryAddPaginatedSplitTableRow(table, rows, rowIndex, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                            consumedRows = true;
                            continue;
                        }

                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return consumedRows;
                        }

                        if (!AddRepeatingTableHeaderRows(table, rows, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                            return consumedRows;
                        }

                        if (context.Y + rowHeight > context.ContentBottom && rows[rowIndex].AllowRowToBreakAcrossPages) {
                            if (TryAddPaginatedSplitTableRow(table, rows, rowIndex, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                                consumedRows = true;
                                continue;
                            }
                        }
                    } else {
                        if (context.IsTargetPage && !context.StoppedForPagination) {
                            AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                            context.StoppedForPagination = true;
                        }

                        return consumedRows;
                    }
                }

                double tableLeft = ResolveTableLeft(table, context.Left, context.ContentWidth, tableWidth);
                if (context.IsTargetPage) {
                    AddTableRow(context, table, rows, rowIndex, tableLeft, context.Y, columnWidths, rowHeights, diagnostics, listMarkers);
                }

                context.Y += rowHeight;
                consumedRows = true;
            }

            context.Y += ParagraphGapPoints;
            return consumedRows;
        }

        private static bool TryAdvanceForTableRowStartPageBreak(
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            int rowIndex,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            int repeatingHeaderRowCount,
            double tableWidth,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            TableRowStartBreakKind breakKind = ResolveTableRowStartBreak(table, rows[rowIndex]);
            if (breakKind == TableRowStartBreakKind.None) {
                return true;
            }

            int previousPageIndex = context.PageIndex;
            if (breakKind == TableRowStartBreakKind.Column) {
                context.AdvanceColumnOrPage();
            } else {
                AdvanceForPageBreakBefore(context);
            }

            if (context.PastTargetPage) {
                return false;
            }

            if (context.PageIndex != previousPageIndex &&
                rowIndex >= repeatingHeaderRowCount &&
                !AddRepeatingTableHeaderRows(table, rows, columnWidths, rowHeights, repeatingHeaderRowCount, tableWidth, context, diagnostics, listMarkers)) {
                return false;
            }

            return true;
        }

        private static bool HasTableRowStartBreak(WordTable table, IReadOnlyList<WordTableRow> rows) =>
            rows.Any(row => ResolveTableRowStartBreak(table, row) != TableRowStartBreakKind.None);

        private static TableRowStartBreakKind ResolveTableRowStartBreak(WordTable table, WordTableRow row) {
            foreach (WordTableCell cell in row.GetCells(readOnly: true)) {
                WordParagraph? firstParagraph = cell.Paragraphs.FirstOrDefault();
                if (firstParagraph != null && ResolvePageBreakBefore(table.Document, firstParagraph._paragraph)) {
                    return TableRowStartBreakKind.Page;
                }

                Paragraph? firstOpenXmlParagraph = cell._tableCell.ChildElements.OfType<Paragraph>().FirstOrDefault();
                WordParagraph? firstRun = firstOpenXmlParagraph == null
                    ? null
                    : WordSection.ConvertParagraphToWordParagraphs(cell.Document, firstOpenXmlParagraph, splitPaginationMarkers: true).FirstOrDefault();
                if (firstRun?.IsPageBreak == true) {
                    return TableRowStartBreakKind.Page;
                }

                if (firstRun?.IsColumnBreak == true) {
                    return TableRowStartBreakKind.Column;
                }
            }

            return TableRowStartBreakKind.None;
        }

        private enum TableRowStartBreakKind {
            None,
            Page,
            Column
        }

        private static bool AddRepeatingTableHeaderRows(
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            int repeatingHeaderRowCount,
            double tableWidth,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            if (repeatingHeaderRowCount == 0) {
                return true;
            }

            double headerHeight = SumHeights(rowHeights, 0, repeatingHeaderRowCount);
            if (headerHeight >= context.ContentHeight) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-table-header-pagination", "Skipped repeating Word table header rows because they do not fit within the page content area.");
                }

                return false;
            }

            double tableLeft = ResolveTableLeft(table, context.Left, context.ContentWidth, tableWidth);
            for (int headerIndex = 0; headerIndex < repeatingHeaderRowCount; headerIndex++) {
                context.ThrowIfCancellationRequested();
                if (context.IsTargetPage) {
                    AddTableRow(context, table, rows, headerIndex, tableLeft, context.Y, columnWidths, rowHeights, diagnostics, listMarkers);
                }

                context.Y += rowHeights[headerIndex];
            }

            return true;
        }

        private static int GetRepeatingHeaderRowCount(IReadOnlyList<WordTableRow> rows) {
            int count = 0;
            for (int i = 0; i < rows.Count; i++) {
                if (!rows[i].RepeatHeaderRowAtTheTopOfEachPage) {
                    break;
                }

                count++;
            }

            return count;
        }

        private static void AddTableRow(
            WordImageFlowContext context,
            WordTable table,
            IReadOnlyList<WordTableRow> rows,
            int rowIndex,
            double tableLeft,
            double rowTop,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            WordTableRow row = rows[rowIndex];
            double cellLeft = tableLeft;
            int columnIndex = 0;

            foreach (WordTableCell cell in row.GetCells(readOnly: true)) {
                context.ThrowIfCancellationRequested();
                int columnSpan = Math.Max(1, cell.ColumnSpan);
                if (cell.HorizontalMerge == MergedCellValues.Continue || cell.VerticalMerge == MergedCellValues.Continue) {
                    columnIndex += columnSpan;
                    cellLeft += SumWidths(columnWidths, columnIndex - columnSpan, columnSpan);
                    continue;
                }

                int rowSpan = Math.Max(1, cell.RowSpan);
                double cellWidth = SumWidths(columnWidths, columnIndex, columnSpan);
                double cellHeight = SumHeights(rowHeights, rowIndex, rowSpan);
                AddTableCell(context, table, cell, rowIndex, columnIndex, rowSpan, columnSpan, rows.Count, columnWidths.Count, cellLeft, rowTop, cellWidth, cellHeight, diagnostics, listMarkers);
                columnIndex += columnSpan;
                cellLeft += cellWidth;
            }
        }

        private static void AddTableCell(
            WordImageFlowContext context,
            WordTable table,
            WordTableCell cell,
            int rowIndex,
            int columnIndex,
            int rowSpan,
            int columnSpan,
            int rowCount,
            int columnCount,
            double left,
            double top,
            double width,
            double height,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            context.ThrowIfCancellationRequested();
            A.ColorScheme? colorScheme = GetDocumentColorScheme(cell.Document);
            OfficeDrawing drawing = context.Drawing;
            drawing.AddBorderBox(
                left,
                top,
                width,
                height,
                ResolveCellFillColor(table, cell, rowIndex, columnIndex, rowCount, columnCount, colorScheme),
                ResolveCellBorders(table, cell, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount, colorScheme));

            double marginLeft = ToPoints(cell.MarginLeftWidth, DefaultCellMarginPoints);
            double marginRight = ToPoints(cell.MarginRightWidth, DefaultCellMarginPoints);
            double marginTop = ToPoints(cell.MarginTopWidth, DefaultCellMarginPoints);
            double marginBottom = ToPoints(cell.MarginBottomWidth, DefaultCellMarginPoints);
            double contentWidth = Math.Max(1D, width - marginLeft - marginRight);
            double contentBottom = top + Math.Max(1D, height - marginBottom);
            double contentLeft = left + marginLeft;
            double contentTop = top + marginTop;
            double textTop = AddTableCellImages(
                cell,
                drawing,
                contentLeft,
                contentTop,
                contentWidth,
                contentBottom,
                diagnostics,
                context.CancellationToken);
            textTop = Math.Max(textTop, AddNestedTables(cell, drawing, contentLeft, textTop, contentWidth, contentBottom, diagnostics, listMarkers, context));

            List<List<WordParagraph>> paragraphRuns = CreateTableCellParagraphRuns(
                cell,
                context.CancellationToken);
            if (paragraphRuns.Count > 1) {
                AddTableCellParagraphFlow(
                    cell,
                    drawing,
                    contentLeft,
                    textTop,
                    contentWidth,
                    Math.Max(1D, contentBottom - textTop),
                    contentBottom,
                    paragraphRuns,
                    listMarkers,
                    colorScheme,
                    diagnostics,
                    context);
                return;
            }

            string text = GetCellText(
                cell,
                context,
                context.CancellationToken);
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            WordParagraph? firstParagraph = cell.Paragraphs.FirstOrDefault(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text));
            OfficeFontInfo font = firstParagraph == null ? OfficeFontInfo.Default : CreateFont(firstParagraph);
            OfficeTextVerticalAlignment verticalAlignment = cell.VerticalAlignment == TableVerticalAlignmentValues.Center
                ? OfficeTextVerticalAlignment.Center
                : cell.VerticalAlignment == TableVerticalAlignmentValues.Bottom
                    ? OfficeTextVerticalAlignment.Bottom
                    : OfficeTextVerticalAlignment.Top;
            var padding = new OfficeTextPadding(marginLeft, marginTop, marginRight, marginBottom);
            bool textFlowAdvanced = textTop > contentTop + 0.000001D;
            double textBoxTop = textFlowAdvanced ? textTop - marginTop : top;
            double textBoxHeight = textFlowAdvanced ? Math.Max(1D, contentBottom - textBoxTop) : height;
            List<OfficeRichTextRun> richRuns = CreateTableCellRichTextRuns(
                cell,
                colorScheme,
                context,
                context.CancellationToken);
            if (ShouldRenderTableCellAsRichText(richRuns)) {
                double maxFontSize = richRuns.Max(run => run.FontSize);
                double lineHeight = Math.Max(maxFontSize * 1.25D, 12D);
                drawing.AddRichText(
                    richRuns,
                    left,
                    textBoxTop,
                    width,
                    textBoxHeight,
                    MapTextAlignment(firstParagraph?.ParagraphAlignment),
                    lineHeight,
                    verticalAlignment,
                    wrapText: true,
                    padding: padding);
                return;
            }

            drawing.AddText(
                text,
                left,
                textBoxTop,
                width,
                textBoxHeight,
                font,
                ResolveParagraphTextColor(firstParagraph, colorScheme),
                MapTextAlignment(firstParagraph?.ParagraphAlignment),
                Math.Max(font.Size * 1.25D, 12D),
                verticalAlignment,
                wrapText: true,
                padding: padding);
        }

        private static double AddNestedTables(
            WordTableCell cell,
            OfficeDrawing drawing,
            double left,
            double top,
            double contentWidth,
            double contentBottom,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            WordImageFlowContext? parentContext = null) {
            List<WordTable> nestedTables = GetDirectNestedTables(cell);
            if (nestedTables.Count == 0) {
                return top;
            }

            WordImageFlowContext nestedContext = CreateFlowContext(
                drawing,
                left,
                top,
                contentWidth,
                contentBottom,
                "unsupported-word-nested-table-overflow",
                "Skipped a nested Word table inside a rendered table cell because it does not fit within the cell content area.",
                resolveDynamicPageFields: parentContext?.ResolveDynamicPageFields ?? false,
                totalPageCount: parentContext?.TotalPageCount ?? 1,
                sectionNumber: parentContext?.SectionNumber ?? 1,
                sectionPageCount: parentContext?.SectionPageCount ?? 1,
                pageNumberValue: parentContext?.PageNumberValue ?? 0,
                pageNumberText: parentContext?.PageNumberText,
                cancellationToken: parentContext?.CancellationToken ?? default);

            for (int i = 0; i < nestedTables.Count; i++) {
                nestedContext.ThrowIfCancellationRequested();
                AddTable(nestedTables[i], nestedContext, diagnostics, listMarkers, allowNestedTable: true);
                if (nestedContext.StoppedForPagination) {
                    break;
                }
            }

            return nestedContext.Y;
        }

        private static bool ShouldRenderTableCellAsRichText(IReadOnlyList<OfficeRichTextRun> richRuns) =>
            richRuns.Count > 1 || richRuns.Any(run => run.BackgroundColor.HasValue);

        private static List<OfficeRichTextRun> CreateTableCellRichTextRuns(
            WordTableCell cell,
            A.ColorScheme? colorScheme,
            WordImageFlowContext? context = null,
            CancellationToken cancellationToken = default) {
            var richRuns = new List<OfficeRichTextRun>();
            foreach (Paragraph paragraph in cell._tableCell.ChildElements.OfType<Paragraph>()) {
                cancellationToken.ThrowIfCancellationRequested();
                List<WordParagraph> paragraphRuns = WordSection.ConvertParagraphToWordParagraphs(
                        cell.Document,
                        paragraph,
                        splitPaginationMarkers: true,
                        cancellationToken)
                    .Where(run => !run.IsPageBreak && !run.IsColumnBreak)
                    .Where(run => !string.IsNullOrEmpty(run.Text))
                    .ToList();
                if (paragraphRuns.Count == 0) {
                    continue;
                }

                if (richRuns.Count > 0) {
                    richRuns.Add(CreateRichTextRun(paragraphRuns[0], colorScheme, Environment.NewLine));
                }

                for (int runIndex = 0; runIndex < paragraphRuns.Count; runIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WordParagraph run = paragraphRuns[runIndex];
                    string text = ResolveImageExportText(run, context);
                    if (!string.IsNullOrEmpty(text)) {
                        richRuns.Add(CreateRichTextRun(run, colorScheme, text));
                    }
                }
            }

            return richRuns;
        }

        private static double AddTableCellImages(
            WordTableCell cell,
            OfficeDrawing drawing,
            double left,
            double top,
            double contentWidth,
            double contentBottom,
            List<OfficeImageExportDiagnostic> diagnostics,
            CancellationToken cancellationToken = default) {
            WordImageFlowContext imageContext = new WordImageFlowContext(
                drawing,
                left,
                top,
                contentWidth,
                contentBottom,
                Array.Empty<WordImageColumnFrame>(),
                "unsupported-word-table-image-overflow",
                "Skipped a Word image inside a rendered table cell because it does not fit within the cell content area.");

            foreach (WordParagraph paragraph in cell.Elements.OfType<WordParagraph>()) {
                cancellationToken.ThrowIfCancellationRequested();
                WordImage? image = paragraph.Image;
                if (image != null) {
                    AddImage(image, imageContext, diagnostics);
                }
            }

            return imageContext.Y;
        }

        private static double[] ResolveColumnWidths(WordTable table, int columnCount, double availableWidth) {
            List<int> gridWidths = table.GridColumnWidth.Where(width => width > 0).ToList();
            if (gridWidths.Count == 0) {
                gridWidths = table.ColumnWidth.Where(width => width > 0).ToList();
            }

            if (gridWidths.Count == 0) {
                return Enumerable.Repeat(availableWidth / columnCount, columnCount).ToArray();
            }

            while (gridWidths.Count < columnCount) {
                gridWidths.Add(gridWidths[gridWidths.Count - 1]);
            }

            if (gridWidths.Count > columnCount) {
                gridWidths = gridWidths.Take(columnCount).ToList();
            }

            double[] widths = gridWidths.Select(width => Math.Max(1D, width / TwipsPerPoint)).ToArray();
            double total = widths.Sum();
            if (total > availableWidth) {
                double scale = availableWidth / total;
                for (int i = 0; i < widths.Length; i++) {
                    widths[i] = Math.Max(1D, widths[i] * scale);
                }
            }

            return widths;
        }

        private static double[] ResolveRowHeights(
            IReadOnlyList<WordTableRow> rows,
            IReadOnlyList<double> columnWidths,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null,
            bool signalNestedMeasurementProgress = false) {
            double[] rowHeights = new double[rows.Count];
            for (int i = 0; i < rows.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                rowHeights[i] = ResolveRowHeight(
                    rows[i],
                    columnWidths,
                    listMarkers,
                    cancellationToken,
                    cancellationCheckpoint);
                if (i == 0 && signalNestedMeasurementProgress) {
                    cancellationCheckpoint?.Invoke(
                        WordImageCancellationCheckpoint.NestedTableMeasurement);
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }

            return rowHeights;
        }

        private static double ResolveRowHeight(
            WordTableRow row,
            IReadOnlyList<double> columnWidths,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            cancellationToken.ThrowIfCancellationRequested();
            TableRowHeight? rowHeight = row._tableRow.TableRowProperties?.OfType<TableRowHeight>().FirstOrDefault();
            double explicitHeight = ResolveTableRowHeightPoints(rowHeight);
            double estimatedHeight = Math.Max(
                MinimumTableRowHeightPoints,
                EstimateRowHeight(
                    row,
                    columnWidths,
                    listMarkers,
                    cancellationToken,
                    cancellationCheckpoint));
            if (explicitHeight <= 0D) {
                return estimatedHeight;
            }

            if (rowHeight?.HeightType?.Value == HeightRuleValues.Exact) {
                return explicitHeight;
            }

            return Math.Max(explicitHeight, estimatedHeight);
        }

        private static double ResolveTableRowHeightPoints(TableRowHeight? rowHeight) {
            if (rowHeight?.Val == null) {
                return 0D;
            }

            return rowHeight.Val.Value / TwipsPerPoint;
        }

        private static double EstimateRowHeight(
            WordTableRow row,
            IReadOnlyList<double> columnWidths,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            cancellationToken.ThrowIfCancellationRequested();
            double height = MinimumTableRowHeightPoints;
            int columnIndex = 0;
            foreach (WordTableCell cell in row.GetCells(readOnly: true)) {
                cancellationToken.ThrowIfCancellationRequested();
                int columnSpan = Math.Max(1, cell.ColumnSpan);
                if (cell.HorizontalMerge == MergedCellValues.Continue || cell.VerticalMerge == MergedCellValues.Continue) {
                    columnIndex += columnSpan;
                    continue;
                }

                double width = Math.Max(1D, SumWidths(columnWidths, columnIndex, columnSpan) - ToPoints(cell.MarginLeftWidth, DefaultCellMarginPoints) - ToPoints(cell.MarginRightWidth, DefaultCellMarginPoints));
                WordParagraph? firstParagraph = cell.Paragraphs.FirstOrDefault(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text));
                OfficeFontInfo font = firstParagraph == null ? OfficeFontInfo.Default : CreateFont(firstParagraph);
                double lineHeight = Math.Max(font.Size * 1.25D, 12D);
                double imageHeight = EstimateCellImageHeight(cell, cancellationToken);
                double nestedTableHeight = EstimateCellNestedTableHeight(
                    cell,
                    width,
                    cancellationToken,
                    cancellationCheckpoint);
                List<List<WordParagraph>> paragraphRuns = CreateTableCellParagraphRuns(
                    cell,
                    cancellationToken);
                double textHeight = EstimateTableCellTextHeight(
                    cell,
                    paragraphRuns,
                    font.Size,
                    width,
                    lineHeight,
                    listMarkers,
                    cancellationToken);
                double stackedHeight = imageHeight + nestedTableHeight + textHeight;
                height = Math.Max(height, stackedHeight + ToPoints(cell.MarginTopWidth, DefaultCellMarginPoints) + ToPoints(cell.MarginBottomWidth, DefaultCellMarginPoints));
                columnIndex += columnSpan;
            }

            return height;
        }

        private static double EstimateCellNestedTableHeight(
            WordTableCell cell,
            double availableWidth,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            double height = 0D;
            List<WordTable> nestedTables = GetDirectNestedTables(cell);
            for (int i = 0; i < nestedTables.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                height += EstimateTableHeight(
                    nestedTables[i],
                    availableWidth,
                    cancellationToken,
                    cancellationCheckpoint,
                    signalNestedMeasurementProgress: true) + ParagraphGapPoints;
            }

            return Math.Max(0D, height - ParagraphGapPoints);
        }

        private static double EstimateTableHeight(
            WordTable table,
            double availableWidth,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null,
            bool signalNestedMeasurementProgress = false) {
            cancellationToken.ThrowIfCancellationRequested();
            List<WordTableRow> rows = table.Rows;
            if (rows.Count == 0) {
                return 0D;
            }

            int columnCount = Math.Max(1, rows.Max(row => row.Cells.Count));
            double[] columnWidths = ResolveColumnWidths(table, columnCount, availableWidth);
            return ResolveRowHeights(
                rows,
                columnWidths,
                listMarkers: null,
                cancellationToken,
                cancellationCheckpoint,
                signalNestedMeasurementProgress).Sum();
        }

        private static double EstimateCellImageHeight(
            WordTableCell cell,
            CancellationToken cancellationToken) {
            double height = 0D;
            foreach (WordParagraph paragraph in cell.Elements.OfType<WordParagraph>()) {
                cancellationToken.ThrowIfCancellationRequested();
                WordImage? image = paragraph.Image;
                if (image == null) {
                    continue;
                }

                height += Helpers.ConvertPixelsToPoints(image.Height ?? 64D) + ParagraphGapPoints;
            }

            return Math.Max(0D, height - ParagraphGapPoints);
        }

        private static double ResolveTableLeft(WordTable table, double contentLeft, double contentWidth, double tableWidth) {
            if (table.Alignment == TableRowAlignmentValues.Center) {
                return contentLeft + Math.Max(0D, (contentWidth - tableWidth) / 2D);
            }

            if (table.Alignment == TableRowAlignmentValues.Right) {
                return contentLeft + Math.Max(0D, contentWidth - tableWidth);
            }

            return contentLeft;
        }

        private static string GetCellText(
            WordTableCell cell,
            WordImageFlowContext? context = null,
            CancellationToken cancellationToken = default) {
            var paragraphs = new List<string>();
            foreach (IReadOnlyList<WordParagraph> runs in CreateTableCellParagraphRuns(
                         cell,
                         cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                var builder = new StringBuilder();
                foreach (WordParagraph run in runs) {
                    cancellationToken.ThrowIfCancellationRequested();
                    builder.Append(ResolveImageExportText(run, context));
                }
                string text = builder.ToString();
                if (!string.IsNullOrWhiteSpace(text)) {
                    paragraphs.Add(text);
                }
            }
            return string.Join("\n", paragraphs);
        }

        private static List<WordTable> GetDirectNestedTables(WordTableCell cell) =>
            cell._tableCell.ChildElements
                .OfType<Table>()
                .Select(table => new WordTable(cell.Document, table, initializeChildren: false))
                .ToList();

        private static OfficeColor ResolveCellFillColor(WordTable table, WordTableCell cell, int rowIndex, int columnIndex, int rowCount, int columnCount, A.ColorScheme? colorScheme) {
            Shading? shading = cell._tableCellProperties?.Shading;
            if (TryResolveShadingFillColor(shading, colorScheme, out OfficeColor directFill)) {
                return directFill;
            }

            if (cell.ShadingFillColor.HasValue) {
                return cell.ShadingFillColor.Value;
            }

            foreach (TableStyleProperties properties in EnumerateApplicableTableConditionalStyleProperties(table, rowIndex, columnIndex, rowCount, columnCount)) {
                Shading? conditionalShading = properties.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.Shading;
                if (TryResolveShadingFillColor(conditionalShading, colorScheme, out OfficeColor conditionalFill)) {
                    return conditionalFill;
                }
            }

            foreach (StyleTableProperties properties in EnumerateTableStyleProperties(table)) {
                if (TryResolveShadingFillColor(properties.Shading, colorScheme, out OfficeColor styleFill)) {
                    return styleFill;
                }
            }

            return OfficeColor.White;
        }

        private static bool TryResolveShadingFillColor(Shading? shading, A.ColorScheme? colorScheme, out OfficeColor fillColor) {
            fillColor = OfficeColor.White;
            string? resolvedThemeColor = ResolveThemeColor(
                GetWordAttribute(shading, "themeFill"),
                GetWordAttribute(shading, "themeFillTint"),
                GetWordAttribute(shading, "themeFillShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeColor, out OfficeColor themeFill)) {
                fillColor = themeFill;
                return true;
            }

            return TryParseOfficeColor(shading?.Fill?.Value, out fillColor);
        }

        private static OfficeBorderBox ResolveCellBorders(WordTable table, WordTableCell cell, int rowIndex, int columnIndex, int rowSpan, int columnSpan, int rowCount, int columnCount, A.ColorScheme? colorScheme) {
            TableCellBorders? borders = cell._tableCellProperties?.TableCellBorders;
            if (borders != null) {
                return new OfficeBorderBox(
                    ResolveCellBorderSide(cell.Borders.LeftStyle, cell.Borders.LeftColorHex, cell.Borders.LeftSize, borders.LeftBorder, colorScheme),
                    ResolveCellBorderSide(cell.Borders.TopStyle, cell.Borders.TopColorHex, cell.Borders.TopSize, borders.TopBorder, colorScheme),
                    ResolveCellBorderSide(cell.Borders.RightStyle, cell.Borders.RightColorHex, cell.Borders.RightSize, borders.RightBorder, colorScheme),
                    ResolveCellBorderSide(cell.Borders.BottomStyle, cell.Borders.BottomColorHex, cell.Borders.BottomSize, borders.BottomBorder, colorScheme),
                    ResolveOptionalCellBorderSide(cell.Borders.TopLeftToBottomRightStyle, cell.Borders.TopLeftToBottomRightColorHex, cell.Borders.TopLeftToBottomRightSize, borders.TopLeftToBottomRightCellBorder, colorScheme),
                    ResolveOptionalCellBorderSide(cell.Borders.TopRightToBottomLeftStyle, cell.Borders.TopRightToBottomLeftColorHex, cell.Borders.TopRightToBottomLeftSize, borders.TopRightToBottomLeftCellBorder, colorScheme));
            }

            List<TableCellBorders> conditionalBorders = EnumerateApplicableTableConditionalStyleProperties(table, rowIndex, columnIndex, rowCount, columnCount)
                .Select(properties => properties.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.TableCellBorders)
                .Where(tableCellBorders => tableCellBorders != null)
                .Select(tableCellBorders => tableCellBorders!)
                .ToList();
            List<TableBorders> styleBorders = EnumerateTableStyleProperties(table)
                .Select(properties => properties.TableBorders)
                .Where(tableBorders => tableBorders != null)
                .Select(tableBorders => tableBorders!)
                .ToList();
            if (conditionalBorders.Count > 0 || styleBorders.Count > 0) {
                return new OfficeBorderBox(
                    ResolveInheritedTableBorderSide(conditionalBorders, styleBorders, TableCellBorderEdge.Left, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount, colorScheme),
                    ResolveInheritedTableBorderSide(conditionalBorders, styleBorders, TableCellBorderEdge.Top, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount, colorScheme),
                    ResolveInheritedTableBorderSide(conditionalBorders, styleBorders, TableCellBorderEdge.Right, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount, colorScheme),
                    ResolveInheritedTableBorderSide(conditionalBorders, styleBorders, TableCellBorderEdge.Bottom, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount, colorScheme),
                    null,
                    null);
            }

            return new OfficeBorderBox(
                ResolveCellBorderSide(cell.Borders.LeftStyle, cell.Borders.LeftColorHex, cell.Borders.LeftSize, borders?.LeftBorder, colorScheme),
                ResolveCellBorderSide(cell.Borders.TopStyle, cell.Borders.TopColorHex, cell.Borders.TopSize, borders?.TopBorder, colorScheme),
                ResolveCellBorderSide(cell.Borders.RightStyle, cell.Borders.RightColorHex, cell.Borders.RightSize, borders?.RightBorder, colorScheme),
                ResolveCellBorderSide(cell.Borders.BottomStyle, cell.Borders.BottomColorHex, cell.Borders.BottomSize, borders?.BottomBorder, colorScheme),
                ResolveOptionalCellBorderSide(cell.Borders.TopLeftToBottomRightStyle, cell.Borders.TopLeftToBottomRightColorHex, cell.Borders.TopLeftToBottomRightSize, borders?.TopLeftToBottomRightCellBorder, colorScheme),
                ResolveOptionalCellBorderSide(cell.Borders.TopRightToBottomLeftStyle, cell.Borders.TopRightToBottomLeftColorHex, cell.Borders.TopRightToBottomLeftSize, borders?.TopRightToBottomLeftCellBorder, colorScheme));
        }

        private enum TableCellBorderEdge {
            Left,
            Top,
            Right,
            Bottom
        }

        private static OfficeBorderSide? ResolveInheritedTableBorderSide(
            IReadOnlyList<TableCellBorders> conditionalBorders,
            IReadOnlyList<TableBorders> styleBorders,
            TableCellBorderEdge edge,
            int rowIndex,
            int columnIndex,
            int rowSpan,
            int columnSpan,
            int rowCount,
            int columnCount,
            A.ColorScheme? colorScheme) {
            foreach (TableCellBorders borders in conditionalBorders) {
                OpenXmlElement? source = SelectCellBorderSource(borders, edge, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount);
                if (source != null) {
                    return ResolveTableBorderSide(source as BorderType, colorScheme);
                }
            }

            foreach (TableBorders borders in styleBorders) {
                OpenXmlElement? source = SelectTableBorderSource(borders, edge, rowIndex, columnIndex, rowSpan, columnSpan, rowCount, columnCount);
                if (source != null) {
                    return ResolveTableBorderSide(source as BorderType, colorScheme);
                }
            }

            return null;
        }

        private static OpenXmlElement? SelectCellBorderSource(TableCellBorders borders, TableCellBorderEdge edge, int rowIndex, int columnIndex, int rowSpan, int columnSpan, int rowCount, int columnCount) =>
            edge switch {
                TableCellBorderEdge.Left => (OpenXmlElement?)borders.LeftBorder ?? (columnIndex == 0 ? null : borders.InsideVerticalBorder),
                TableCellBorderEdge.Top => (OpenXmlElement?)borders.TopBorder ?? (rowIndex == 0 ? null : borders.InsideHorizontalBorder),
                TableCellBorderEdge.Right => (OpenXmlElement?)borders.RightBorder ?? (columnIndex + columnSpan >= columnCount ? null : borders.InsideVerticalBorder),
                TableCellBorderEdge.Bottom => (OpenXmlElement?)borders.BottomBorder ?? (rowIndex + rowSpan >= rowCount ? null : borders.InsideHorizontalBorder),
                _ => null
            };

        private static OpenXmlElement? SelectTableBorderSource(TableBorders borders, TableCellBorderEdge edge, int rowIndex, int columnIndex, int rowSpan, int columnSpan, int rowCount, int columnCount) =>
            edge switch {
                TableCellBorderEdge.Left => columnIndex == 0 ? borders.LeftBorder : borders.InsideVerticalBorder,
                TableCellBorderEdge.Top => rowIndex == 0 ? borders.TopBorder : borders.InsideHorizontalBorder,
                TableCellBorderEdge.Right => columnIndex + columnSpan >= columnCount ? borders.RightBorder : borders.InsideVerticalBorder,
                TableCellBorderEdge.Bottom => rowIndex + rowSpan >= rowCount ? borders.BottomBorder : borders.InsideHorizontalBorder,
                _ => null
            };

        private static OfficeBorderSide? ResolveTableBorderSide(BorderType? source, A.ColorScheme? colorScheme) {
            if (source == null || source.Val == null || source.Val.Value == BorderValues.Nil || source.Val.Value == BorderValues.None) {
                return null;
            }

            return CreateCellBorderSide(source.Val.Value, source.Color?.Value, source.Size, source, colorScheme);
        }

        private static OfficeBorderSide? ResolveCellBorderSide(BorderValues? style, string? color, UInt32Value? size, OpenXmlElement? source, A.ColorScheme? colorScheme) {
            if (style == BorderValues.Nil || style == BorderValues.None) {
                return null;
            }

            return CreateCellBorderSide(style, color, size, source, colorScheme);
        }

        private static OfficeBorderSide? ResolveOptionalCellBorderSide(BorderValues? style, string? color, UInt32Value? size, OpenXmlElement? source, A.ColorScheme? colorScheme) {
            if (!style.HasValue || style == BorderValues.Nil || style == BorderValues.None) {
                return null;
            }

            return CreateCellBorderSide(style, color, size, source, colorScheme);
        }

        private static OfficeBorderSide CreateCellBorderSide(BorderValues? style, string? color, UInt32Value? size, OpenXmlElement? source, A.ColorScheme? colorScheme) {
            double width = ResolveCellBorderWidth(size);
            return new OfficeBorderSide(
                ResolveCellBorderColor(color, source, colorScheme),
                width,
                MapBorderDashStyle(style),
                style == BorderValues.Double ? OfficeBorderLineKind.Double : OfficeBorderLineKind.Single,
                style == BorderValues.Double ? Math.Max(1.5D, width * 3D) : 0D);
        }

        private static OfficeColor ResolveCellBorderColor(string? color, OpenXmlElement? source, A.ColorScheme? colorScheme) {
            string? resolvedThemeColor = ResolveThemeColor(
                GetWordAttribute(source, "themeColor"),
                GetWordAttribute(source, "themeTint"),
                GetWordAttribute(source, "themeShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeColor, out OfficeColor themeColor)) {
                return themeColor;
            }

            if (!string.IsNullOrWhiteSpace(color) && !string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) {
                try {
                    return Helpers.ParseColor(color!);
                } catch (ArgumentException) {
                    return OfficeColor.LightGray;
                } catch (FormatException) {
                    return OfficeColor.LightGray;
                }
            }

            return OfficeColor.LightGray;
        }

        private static double ResolveCellBorderWidth(UInt32Value? borderSize) {
            uint? size = borderSize?.Value;
            if (!size.HasValue || size.Value == 0U) {
                return 0.75D;
            }

            return Math.Max(0.5D, size.Value / 8D);
        }

        private static OfficeStrokeDashStyle MapBorderDashStyle(BorderValues? style) {
            if (style == BorderValues.Dashed || style == BorderValues.DashSmallGap) {
                return OfficeStrokeDashStyle.Dash;
            }

            if (style == BorderValues.Dotted) {
                return OfficeStrokeDashStyle.Dot;
            }

            if (style == BorderValues.DotDash) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (style == BorderValues.DotDotDash) {
                return OfficeStrokeDashStyle.DashDotDot;
            }

            return OfficeStrokeDashStyle.Solid;
        }

        private static double SumWidths(IReadOnlyList<double> widths, int start, int count) {
            double sum = 0D;
            for (int i = 0; i < count && start + i < widths.Count; i++) {
                sum += widths[start + i];
            }

            return Math.Max(1D, sum);
        }

        private static double SumHeights(IReadOnlyList<double> heights, int start, int count) {
            double sum = 0D;
            for (int i = 0; i < count && start + i < heights.Count; i++) {
                sum += heights[start + i];
            }

            return Math.Max(1D, sum);
        }
    }
}
