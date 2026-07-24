using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int NativeOfficeImoScaffoldCellWidthTwips = 2400;
        private const double NativeAutoFitGridMinimumScale = 0.8D;

        private static void RenderNativeTable(INativePdfFlow pdf, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            RecordNativeBodyTableDiagnostics(table, options, "body table");

            TableLayout layout = TableLayoutCache.GetLayout(table);
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
            var rows = new List<PdfCore.PdfTableCell[]>();
            var cellFills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            var cellBorders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            var cellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
            var cellAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            var cellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
            var horizontalAlignments = CreateNativeTableHorizontalAlignments(layout);
            var verticalAlignments = CreateNativeTableVerticalAlignments(layout);
            int tableColumnCount = GetNativeTableColumnCount(layout);
            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, layout.Rows.Count);
            int visualHeaderRowCount = GetNativeTableVisualHeaderRowCount(table, layout.Rows.Count, repeatedHeaderRowCount);
            int footerStartRowIndex = table.ConditionalFormattingLastRow == true && layout.Rows.Count > visualHeaderRowCount
                ? layout.Rows.Count - 1
                : layout.Rows.Count;
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                var nativeCells = new List<PdfCore.PdfTableCell>();
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                AddNativeTableGridBeforePlaceholders(nativeCells, logicalColumnIndex);
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    WordTableCell cell = row[columnIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    NativeTableStyleDefaults cellStyleDefaults = GetNativeTableCellStyleDefaults(
                        table,
                        tableStyleDefaults,
                        rowIndex,
                        logicalColumnIndex,
                        columnSpan,
                        tableColumnCount,
                        visualHeaderRowCount,
                        footerStartRowIndex);
                    NativeCellText cellText = CreateNativeCellText(cell, footnoteNumbersById, nativeDefaults, cellStyleDefaults, nativeFontMap);
                    IReadOnlyList<PdfCore.PdfTableCellCheckBox> checkBoxes = CreateNativeTableCellCheckBoxes(cell);
                    IReadOnlyList<PdfCore.PdfTableCellFormField> formFields = CreateNativeTableCellFormFields(cell);
                    IReadOnlyList<PdfCore.PdfTableCellImage> images = CreateNativeTableCellImages(cell);
                    (string? LinkUri, string? LinkContents) link = GetNativeCellLink(cell);
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellText.Runs,
                        cellText.Paragraphs,
                        columnSpan,
                        link.LinkUri,
                        link.LinkContents,
                        rowSpan,
                        checkBoxes.Count == 0 ? null : checkBoxes,
                        formFields.Count == 0 ? null : formFields,
                        images.Count == 0 ? null : images,
                        noWrap: !cell.WrapText));

                    PdfCore.PdfColor? fill = ParseNativeColor(cell.ShadingFillColorHex) ?? tableStyleDefaults.CellFill;
                    if (fill.HasValue) {
                        cellFills[(rowIndex, logicalColumnIndex)] = fill.Value;
                    }

                    PdfCore.PdfCellBorder? border = CreateNativeTableCellBorder(cell.Borders);
                    if (border != null) {
                        cellBorders[(rowIndex, logicalColumnIndex)] = border;
                    }

                    PdfCore.PdfCellPadding? padding = CreateNativeTableCellPadding(cell);
                    if (padding != null) {
                        cellPaddings[(rowIndex, logicalColumnIndex)] = padding;
                    }

                    PdfCore.PdfColumnAlign cellAlignment = GetNativeCellHorizontalAlignment(cell);
                    if (cellAlignment != PdfCore.PdfColumnAlign.Left) {
                        cellAlignments[(rowIndex, logicalColumnIndex)] = cellAlignment;
                    }

                    PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = ResolveNativeTableCellVerticalAlignment(cell, cellStyleDefaults);
                    if (cellVerticalAlignment.HasValue) {
                        cellVerticalAlignments[(rowIndex, logicalColumnIndex)] = cellVerticalAlignment.Value;
                    }

                    logicalColumnIndex += columnSpan;
                }

                AddNativeTableGridAfterPlaceholders(nativeCells, GetNativeTableRowTrailingColumnCount(layout, rowIndex));
                rows.Add(nativeCells.ToArray());
            }

            if (rows.Count == 0) {
                return;
            }

            PdfCore.PdfTableStyle style = CreateNativeTableStyle(
                table,
                rows.Count,
                options,
                contentWidth,
                nativeDefaults,
                tableStyleDefaults,
                layout,
                nativeFontMap);
            if (cellFills.Count > 0) {
                if (style.CellFills == null) {
                    style.CellFills = cellFills;
                } else {
                    foreach (var cellFill in cellFills) {
                        style.CellFills[cellFill.Key] = cellFill.Value;
                    }
                }
            }

            if (cellBorders.Count > 0) {
                if (style.CellBorders == null) {
                    style.CellBorders = cellBorders;
                } else {
                    foreach (var cellBorder in cellBorders) {
                        style.CellBorders[cellBorder.Key] = cellBorder.Value;
                    }
                }
            }

            if (cellPaddings.Count > 0) {
                if (style.CellPaddings == null) {
                    style.CellPaddings = cellPaddings;
                } else {
                    var mergedPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(style.CellPaddings);
                    foreach (var cellPadding in cellPaddings) {
                        mergedPaddings[cellPadding.Key] = MergeNativeCellPadding(
                            mergedPaddings.TryGetValue(cellPadding.Key, out PdfCore.PdfCellPadding? existing) ? existing : null,
                            cellPadding.Value)!;
                    }

                    style.CellPaddings = mergedPaddings;
                }
            }

            if (cellAlignments.Count > 0) {
                style.CellAlignments = cellAlignments;
            }

            if (cellVerticalAlignments.Count > 0) {
                if (style.CellVerticalAlignments == null) {
                    style.CellVerticalAlignments = cellVerticalAlignments;
                } else {
                    var mergedVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>(style.CellVerticalAlignments);
                    foreach (var cellVerticalAlignment in cellVerticalAlignments) {
                        mergedVerticalAlignments[cellVerticalAlignment.Key] = cellVerticalAlignment.Value;
                    }

                    style.CellVerticalAlignments = mergedVerticalAlignments;
                }
            }

            ApplyNativeColumnWidths(table, layout, style, contentWidth);

            if (horizontalAlignments != null) {
                style.Alignments = horizontalAlignments;
            }

            if (verticalAlignments != null) {
                style.VerticalAlignments = verticalAlignments;
            }

            pdf.Table(rows, MapNativeTableAlignment(ResolveNativeTableAlignment(table, tableStyleDefaults)), style);
        }

        private static void ApplyNativeColumnWidths(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            List<double>? columnWidthWeights = CreateNativeColumnWidthWeights(layout);
            if (columnWidthWeights != null) {
                style.ColumnWidthPoints = null;
                style.ColumnWidthWeights = columnWidthWeights;
                return;
            }

            if (style.AutoFitColumns) {
                if (layout.ColumnWidths.Length > 0 && layout.ColumnWidths.All(width => width > 0)) {
                    ApplyNativeAutoFitGridMinimums(layout, style, contentWidth);
                }
                if (layout.ColumnWidths.Length > 0 &&
                    layout.ColumnWidths.All(width => width > 0) &&
                    layout.ColumnWidths.Skip(1).Any(width => Math.Abs(width - layout.ColumnWidths[0]) > 0.01F)) {
                    style.ColumnWidthWeights = layout.ColumnWidths.Select(width => (double)width).ToList();
                }
                return;
            }

            style.ColumnWidthPoints = CreateNativeColumnWidthPoints(layout, style);
        }

        private static void ApplyNativeAutoFitGridMinimums(TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            double tableWidth = style.MaxWidth ?? style.PreferredWidth ?? contentWidth ?? 0D;
            double gridWidth = layout.ColumnWidths.Sum(width => (double)width);
            if (tableWidth <= 0D ||
                gridWidth <= 0D ||
                double.IsNaN(tableWidth) ||
                double.IsInfinity(tableWidth)) {
                return;
            }

            style.ColumnMinWidthPoints = layout.ColumnWidths
                .Select(width => (double?)(tableWidth * width / gridWidth * NativeAutoFitGridMinimumScale))
                .ToList();
        }

        private static List<double>? CreateNativeColumnWidthWeights(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var weights = new double[columnCount];
            var hasPercentWidth = new bool[columnCount];
            bool hasAnyPercentWidth = false;
            foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                double? percent = GetNativeTableCellPreferredWidthPercent(cell.Cell);
                if (!percent.HasValue) {
                    continue;
                }

                double columnWeight = percent.Value / cell.ColumnSpan;
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < weights.Length; columnIndex++) {
                    if (!hasPercentWidth[columnIndex] || columnWeight > weights[columnIndex]) {
                        weights[columnIndex] = columnWeight;
                        hasPercentWidth[columnIndex] = true;
                    }
                }

                hasAnyPercentWidth = true;
            }

            if (!hasAnyPercentWidth) {
                return null;
            }

            double fallbackWeight = 0D;
            int weightedColumnCount = 0;
            for (int columnIndex = 0; columnIndex < weights.Length; columnIndex++) {
                if (hasPercentWidth[columnIndex]) {
                    fallbackWeight += weights[columnIndex];
                    weightedColumnCount++;
                }
            }

            fallbackWeight = weightedColumnCount == 0 ? 1D : fallbackWeight / weightedColumnCount;
            for (int columnIndex = 0; columnIndex < weights.Length; columnIndex++) {
                if (!hasPercentWidth[columnIndex]) {
                    weights[columnIndex] = fallbackWeight;
                }
            }

            return weights.ToList();
        }

        private static List<double?>? CreateNativeColumnWidthPoints(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || layout.ColumnWidths.Length == 0 || !layout.ColumnWidths.All(width => width > 0)) {
                return null;
            }

            var widths = layout.ColumnWidths.Select(width => (double)width).ToList();
            double totalWidth = widths.Sum();
            if (style.MaxWidth.HasValue && totalWidth > style.MaxWidth.Value + 0.001D) {
                double scale = style.MaxWidth.Value / totalWidth;
                for (int i = 0; i < widths.Count; i++) {
                    widths[i] *= scale;
                }
            }

            return widths.Select(width => (double?)width).ToList();
        }

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options) =>
            CreateNativeTableStyle(table, rowCount, options, null);

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options, double? contentWidth) =>
            CreateNativeTableStyle(table, rowCount, options, contentWidth, NativeDocumentDefaults.WordDefault);

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults) {
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
            return CreateNativeTableStyle(table, rowCount, options, contentWidth, nativeDefaults, tableStyleDefaults, TableLayoutCache.GetLayout(table));
        }

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(
            WordTable table,
            int rowCount,
            PdfSaveOptions? options,
            double? contentWidth,
            NativeDocumentDefaults nativeDefaults,
            NativeTableStyleDefaults tableStyleDefaults,
            TableLayout layout,
            NativeFontMap? nativeFontMap = null) {
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            PdfCore.PdfTableStyle? wordStyle = ResolveNativeWordTableStyle(table, hasExplicitDefaultTableStyle);
            bool usesConfiguredDefaultStyle = wordStyle == null && hasExplicitDefaultTableStyle;
            PdfCore.PdfTableStyle style = wordStyle ?? CreateNativeDefaultTableStyle(options);
            if (!usesConfiguredDefaultStyle) {
                style.FontSize ??= nativeDefaults.FontSize;
                double? tableParagraphLineHeight = ShouldApplyNativeTableStyleParagraphLineHeight(table)
                    ? ResolveNativeTableStyleParagraphLineHeight(
                        tableStyleDefaults,
                        style.FontSize ?? nativeDefaults.FontSize,
                        nativeDefaults.FontFamily,
                        nativeFontMap)
                    : null;
                style.LineHeight ??= tableParagraphLineHeight ?? nativeDefaults.ParagraphLineHeight;
            }

            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, rowCount);
            style.HeaderRowCount = GetNativeTableVisualHeaderRowCount(table, rowCount, repeatedHeaderRowCount);
            style.RepeatHeaderRowCount = repeatedHeaderRowCount;
            if (repeatedHeaderRowCount > 0) {
                style.PageContinuationSpacingBefore = Math.Max(style.PageContinuationSpacingBefore, NativeTablePageContinuationSpacingBefore);
            }

            if (options?.DefaultTableBorders == true && style.BorderColor == null) {
                style.BorderColor = PdfCore.PdfColor.LightGray;
            }

            ApplyNativeTableAccessibilityText(table, style);
            ApplyNativeTableBorders(table, style, tableStyleDefaults);
            ApplyNativeTableDefaultCellMargins(
                table,
                style,
                usesConfiguredDefaultStyle,
                ShouldApplyNativeTableStyleCellPadding(table) ? tableStyleDefaults : NativeTableStyleDefaults.Empty);
            ApplyNativeTableConditionalStyles(table, style, tableStyleDefaults, rowCount, layout);
            ApplyNativeTableBandingStyles(table, layout, style, tableStyleDefaults);
            ApplyNativeTableConditionalColumnFills(table, layout, tableStyleDefaults, style);
            ApplyNativeTableConditionalBorders(table, layout, tableStyleDefaults, style);
            ApplyNativeTableConditionalPaddings(table, layout, tableStyleDefaults, style);
            ApplyNativeTableLayoutOptions(table, layout, style, contentWidth, tableStyleDefaults);
            ApplyNativeTableRowOptions(table, style);
            SuppressNativeTableRoleBoundariesCrossedByRowSpans(style, layout);
            return style;
        }

        private static void ApplyNativeTableAccessibilityText(WordTable table, PdfCore.PdfTableStyle style) {
            string? alternativeText = FirstNonWhiteSpace(table.Description, table.Title);
            if (!string.IsNullOrWhiteSpace(alternativeText)) {
                style.AlternativeText = alternativeText;
            }
        }

        private static double? ResolveNativeTableStyleParagraphLineHeight(
            NativeTableStyleDefaults tableStyleDefaults,
            double fontSize,
            string? documentFontFamily,
            NativeFontMap? nativeFontMap = null) {
            if (tableStyleDefaults.ParagraphLineSpacingPoints.HasValue && fontSize > 0D) {
                return ResolveNativeLineSpacingHeight(
                    tableStyleDefaults.ParagraphLineSpacingPoints.Value,
                    tableStyleDefaults.ParagraphLineSpacingRule,
                    fontSize,
                    ResolveNativeWordSingleLineHeight(
                        nativeFontMap,
                        tableStyleDefaults.RunStyle.FontFamily,
                        documentFontFamily));
            }

            return tableStyleDefaults.ParagraphLineHeight;
        }

        private static PdfCore.PdfTableStyle CreateNativeDefaultTableStyle(PdfSaveOptions? options) {
            PdfCore.PdfTableStyle? configuredStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true
                ? options.PdfOptions.DefaultTableStyle
                : null;
            if (configuredStyle != null) {
                return configuredStyle.Clone();
            }

            return new PdfCore.PdfTableStyle {
                RowStripeFill = null
            };
        }

        private static void ApplyNativeTableConditionalStyles(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults, int rowCount, TableLayout layout) {
            ApplyNativeFirstRowConditionalStyle(table, style, tableStyleDefaults);
            ApplyNativeLastRowConditionalStyle(table, style, tableStyleDefaults, rowCount, layout);
        }

        private static void ApplyNativeFirstRowConditionalStyle(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults) {
            if (table.ConditionalFormattingFirstRow != true || style.HeaderRowCount <= 0) {
                return;
            }

            ApplyNativeHeaderConditionalStyle(style, tableStyleDefaults.FirstRowStyle);
        }

        private static void ApplyNativeLastRowConditionalStyle(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults, int rowCount, TableLayout layout) {
            if (table.ConditionalFormattingLastRow != true || rowCount <= style.HeaderRowCount) {
                return;
            }

            if (!HasNativeCellSpanningRowBoundary(layout, rowCount - 1)) {
                style.FooterRowCount = 1;
            }
            ApplyNativeFooterConditionalStyle(style, tableStyleDefaults.LastRowStyle);
        }

        private static void SuppressNativeTableRoleBoundariesCrossedByRowSpans(PdfCore.PdfTableStyle style, TableLayout layout) {
            if (style.HeaderRowCount > 0 && HasNativeCellSpanningRowBoundary(layout, style.HeaderRowCount)) {
                // A repeated or semantic header cannot contain only part of a vertically merged Word cell.
                // Preserve fill formatting that otherwise depends on the header role before clearing it.
                ProjectNativeHeaderFillToCells(style, layout, style.HeaderRowCount);
                style.HeaderRowCount = 0;
                style.RepeatHeaderRowCount = 0;
            }
        }

        private static void ProjectNativeHeaderFillToCells(PdfCore.PdfTableStyle style, TableLayout layout, int headerRowCount) {
            if (!style.HeaderFill.HasValue || headerRowCount <= 0) {
                return;
            }

            var cellFills = style.CellFills == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfColor>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfColor>(style.CellFills);
            int projectedRowCount = System.Math.Min(headerRowCount, layout.Rows.Count);
            for (int rowIndex = 0; rowIndex < projectedRowCount; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                foreach (WordTableCell cell in row) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (!IsNativeVerticalMergeContinuation(cell)) {
                        (int Row, int Column) key = (rowIndex, logicalColumnIndex);
                        if (!cellFills.ContainsKey(key)) {
                            cellFills[key] = style.HeaderFill.Value;
                        }
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            style.CellFills = cellFills;
        }

        private static bool HasNativeCellSpanningRowBoundary(TableLayout layout, int boundaryRowIndex) {
            for (int rowIndex = 0; rowIndex < boundaryRowIndex && rowIndex < layout.Rows.Count; rowIndex++) {
                foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                    if (IsNativeHorizontalMergeContinuation(cell) || IsNativeVerticalMergeContinuation(cell)) {
                        continue;
                    }

                    if (rowIndex + GetNativeCellRowSpan(cell) > boundaryRowIndex) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void ApplyNativeHeaderConditionalStyle(PdfCore.PdfTableStyle style, NativeTableConditionalStyleDefaults conditionalStyle) {
            if (conditionalStyle.CellFill.HasValue) {
                style.HeaderFill = conditionalStyle.CellFill.Value;
            }

            if (conditionalStyle.TextColor.HasValue) {
                style.HeaderTextColor = conditionalStyle.TextColor.Value;
            }

            if (conditionalStyle.FontSize.HasValue) {
                style.HeaderFontSize = conditionalStyle.FontSize.Value;
            }

            if (conditionalStyle.Bold.HasValue) {
                style.HeaderBold = conditionalStyle.Bold.Value;
            }
        }

        private static void ApplyNativeFooterConditionalStyle(PdfCore.PdfTableStyle style, NativeTableConditionalStyleDefaults conditionalStyle) {
            if (conditionalStyle.CellFill.HasValue) {
                style.FooterFill = conditionalStyle.CellFill.Value;
            }

            if (conditionalStyle.TextColor.HasValue) {
                style.FooterTextColor = conditionalStyle.TextColor.Value;
            }

            if (conditionalStyle.FontSize.HasValue) {
                style.FooterFontSize = conditionalStyle.FontSize.Value;
            }

            if (conditionalStyle.Bold.HasValue) {
                style.FooterBold = conditionalStyle.Bold.Value;
            }
        }

        private static void ApplyNativeTableBandingStyles(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults) {
            if (table.ConditionalFormattingNoHorizontalBand != true && tableStyleDefaults.Band1HorizontalStyle.CellFill.HasValue) {
                style.RowStripeFill = tableStyleDefaults.Band1HorizontalStyle.CellFill.Value;
            }

            if (table.ConditionalFormattingNoVerticalBand != true && tableStyleDefaults.Band1VerticalStyle.CellFill.HasValue) {
                ApplyNativeTableVerticalBandingFill(layout, style, tableStyleDefaults.Band1VerticalStyle.CellFill.Value);
            }
        }

        private static void ApplyNativeTableVerticalBandingFill(TableLayout layout, PdfCore.PdfTableStyle style, PdfCore.PdfColor fill) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return;
            }

            var bodyColumnFills = style.BodyColumnFills == null
                ? new List<PdfCore.PdfColor?>(new PdfCore.PdfColor?[columnCount])
                : new List<PdfCore.PdfColor?>(style.BodyColumnFills);
            while (bodyColumnFills.Count < columnCount) {
                bodyColumnFills.Add(null);
            }

            for (int columnIndex = 1; columnIndex < columnCount; columnIndex += 2) {
                bodyColumnFills[columnIndex] = fill;
            }

            style.BodyColumnFills = bodyColumnFills;
        }

        private static void ApplyNativeTableConditionalColumnFills(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, PdfCore.PdfTableStyle style) {
            Dictionary<(int Row, int Column), PdfCore.PdfColor>? cellFills = style.CellFills == null
                ? null
                : new Dictionary<(int Row, int Column), PdfCore.PdfColor>(style.CellFills);
            cellFills ??= new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            int originalCount = cellFills.Count;
            ApplyNativeTableConditionalColumnFills(table, layout, tableStyleDefaults, cellFills);
            if (cellFills.Count != originalCount) {
                style.CellFills = cellFills;
            }
        }

        private static void ApplyNativeTableConditionalColumnFills(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, Dictionary<(int Row, int Column), PdfCore.PdfColor> cellFills) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return;
            }

            PdfCore.PdfColor? firstColumnFill = table.ConditionalFormattingFirstColumn == true
                ? tableStyleDefaults.FirstColumnStyle.CellFill
                : null;
            PdfCore.PdfColor? lastColumnFill = table.ConditionalFormattingLastColumn == true
                ? tableStyleDefaults.LastColumnStyle.CellFill
                : null;
            if (!firstColumnFill.HasValue && !lastColumnFill.HasValue) {
                return;
            }

            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    WordTableCell cell = row[cellIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    (int Row, int Column) key = (rowIndex, logicalColumnIndex);
                    if (firstColumnFill.HasValue && logicalColumnIndex == 0 && !cellFills.ContainsKey(key)) {
                        cellFills[key] = firstColumnFill.Value;
                    }

                    if (lastColumnFill.HasValue && logicalColumnIndex + columnSpan >= columnCount && !cellFills.ContainsKey(key)) {
                        cellFills[key] = lastColumnFill.Value;
                    }

                    logicalColumnIndex += columnSpan;
                }
            }
        }

        private static NativeTableStyleDefaults GetNativeTableCellStyleDefaults(WordTable table, NativeTableStyleDefaults tableStyleDefaults, int rowIndex, int logicalColumnIndex, int columnSpan, int columnCount, int headerRowCount, int footerStartRowIndex) {
            NativeTableStyleDefaults result = tableStyleDefaults;
            if (table.ConditionalFormattingFirstRow == true && rowIndex == 0) {
                result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.FirstRowStyle);
            }

            if (table.ConditionalFormattingLastRow == true && rowIndex >= footerStartRowIndex) {
                result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.LastRowStyle);
            }

            if (rowIndex >= headerRowCount && rowIndex < footerStartRowIndex) {
                int bodyRowIndex = rowIndex - headerRowCount;
                if (table.ConditionalFormattingNoHorizontalBand != true && bodyRowIndex % 2 == 1) {
                    result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.Band1HorizontalStyle);
                }

                if (table.ConditionalFormattingNoVerticalBand != true && logicalColumnIndex % 2 == 1) {
                    result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.Band1VerticalStyle);
                }
            }

            if (table.ConditionalFormattingFirstColumn == true && logicalColumnIndex == 0) {
                result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.FirstColumnStyle);
            }

            if (table.ConditionalFormattingLastColumn == true && columnCount > 0 && logicalColumnIndex + columnSpan >= columnCount) {
                result = ApplyNativeTableConditionalStyleDefaults(result, tableStyleDefaults.LastColumnStyle);
            }

            return result;
        }

        private static NativeTableStyleDefaults ApplyNativeTableConditionalStyleDefaults(NativeTableStyleDefaults tableStyleDefaults, NativeTableConditionalStyleDefaults conditionalStyle) {
            if (!conditionalStyle.TextColor.HasValue &&
                !conditionalStyle.FontSize.HasValue &&
                !conditionalStyle.Bold.HasValue &&
                !conditionalStyle.Italic.HasValue &&
                !conditionalStyle.Underline.HasValue &&
                !conditionalStyle.Strike.HasValue &&
                !conditionalStyle.AllCaps.HasValue &&
                !conditionalStyle.Baseline.HasValue &&
                !conditionalStyle.Highlight.HasValue &&
                !conditionalStyle.CellVerticalAlignment.HasValue &&
                !conditionalStyle.ParagraphLineHeight.HasValue &&
                !conditionalStyle.ParagraphLineSpacingPoints.HasValue &&
                !conditionalStyle.ParagraphLineSpacingRule.HasValue &&
                !conditionalStyle.ParagraphSpacingBefore.HasValue &&
                !conditionalStyle.ParagraphSpacingAfter.HasValue &&
                !conditionalStyle.ParagraphAlignment.HasValue &&
                !conditionalStyle.ParagraphLeftIndent.HasValue &&
                !conditionalStyle.ParagraphRightIndent.HasValue &&
                !conditionalStyle.ParagraphFirstLineIndent.HasValue) {
                return tableStyleDefaults;
            }

            NativeTableRunStyleDefaults runStyle = tableStyleDefaults.RunStyle;
            return tableStyleDefaults with {
                CellVerticalAlignment = conditionalStyle.CellVerticalAlignment ?? tableStyleDefaults.CellVerticalAlignment,
                ParagraphLineHeight = conditionalStyle.ParagraphLineHeight ?? tableStyleDefaults.ParagraphLineHeight,
                ParagraphLineSpacingPoints = conditionalStyle.ParagraphLineSpacingPoints ?? tableStyleDefaults.ParagraphLineSpacingPoints,
                ParagraphLineSpacingRule = conditionalStyle.ParagraphLineSpacingRule ?? tableStyleDefaults.ParagraphLineSpacingRule,
                ParagraphSpacingBefore = conditionalStyle.ParagraphSpacingBefore ?? tableStyleDefaults.ParagraphSpacingBefore,
                ParagraphSpacingAfter = conditionalStyle.ParagraphSpacingAfter ?? tableStyleDefaults.ParagraphSpacingAfter,
                ParagraphAlignment = conditionalStyle.ParagraphAlignment ?? tableStyleDefaults.ParagraphAlignment,
                ParagraphLeftIndent = conditionalStyle.ParagraphLeftIndent ?? tableStyleDefaults.ParagraphLeftIndent,
                ParagraphRightIndent = conditionalStyle.ParagraphRightIndent ?? tableStyleDefaults.ParagraphRightIndent,
                ParagraphFirstLineIndent = conditionalStyle.ParagraphFirstLineIndent ?? tableStyleDefaults.ParagraphFirstLineIndent,
                RunStyle = runStyle with {
                    FontSize = conditionalStyle.FontSize ?? runStyle.FontSize,
                    Bold = conditionalStyle.Bold ?? runStyle.Bold,
                    Italic = conditionalStyle.Italic ?? runStyle.Italic,
                    Underline = conditionalStyle.Underline ?? runStyle.Underline,
                    Strike = conditionalStyle.Strike ?? runStyle.Strike,
                    AllCaps = conditionalStyle.AllCaps ?? runStyle.AllCaps,
                    Baseline = conditionalStyle.Baseline ?? runStyle.Baseline,
                    Color = conditionalStyle.TextColor ?? runStyle.Color,
                    Highlight = conditionalStyle.Highlight ?? runStyle.Highlight
                }
            };
        }

        private static PdfCore.PdfCellVerticalAlign? ResolveNativeTableCellVerticalAlignment(WordTableCell cell, NativeTableStyleDefaults cellStyleDefaults) {
            PdfCore.PdfCellVerticalAlign? directAlignment = MapNativeNullableCellVerticalAlign(cell.VerticalAlignment);
            if (directAlignment.HasValue) {
                return directAlignment.Value;
            }

            PdfCore.PdfCellVerticalAlign? styleAlignment = cellStyleDefaults.CellVerticalAlignment;
            return styleAlignment.HasValue && styleAlignment.Value != PdfCore.PdfCellVerticalAlign.Top
                ? styleAlignment.Value
                : null;
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableProperties? properties = table._tableProperties;
            if (ShouldUseNativeAutoFitTableLayout(table, properties, tableStyleDefaults)) {
                style.AutoFitColumns = true;
            }

            double? cellSpacing = GetNativeTableCellSpacing(properties?.TableCellSpacing) ?? tableStyleDefaults.CellSpacing;
            if (cellSpacing.HasValue) {
                style.CellSpacing = cellSpacing.Value;
            }

            double? maxWidth = GetNativeTablePreferredWidth(properties?.TableWidth, contentWidth) ??
                GetNativeTablePreferredWidth(tableStyleDefaults.PreferredWidth, contentWidth);
            if (maxWidth.HasValue) {
                style.MaxWidth = maxWidth.Value;
                style.PreserveWidth = true;
            } else {
                double? preferredWidth = GetNativeAutoFitGridPreferredWidth(properties, layout, contentWidth, style.CellSpacing);
                if (preferredWidth.HasValue) {
                    style.PreferredWidth = preferredWidth.Value;
                    style.PreserveWidth = true;
                }
            }

            double? leftIndent = GetNativeTableHorizontalPositionIndent(properties?.TablePositionProperties) ??
                GetNativeTableLeftIndent(properties?.TableIndentation) ??
                tableStyleDefaults.LeftIndent;
            if (leftIndent.HasValue) {
                style.LeftIndent = leftIndent.Value;
            }

        }

        private static double? GetNativeAutoFitGridPreferredWidth(W.TableProperties? properties, TableLayout layout, double? contentWidth, double cellSpacing) {
            if (!IsNativeTableAutoFitToContents(properties) ||
                layout.ColumnWidths.Length == 0 ||
                !layout.ColumnWidths.All(width => width > 0F)) {
                return null;
            }

            double gridWidth = layout.ColumnWidths.Sum(width => (double)width) +
                Math.Max(0, layout.ColumnWidths.Length - 1) * cellSpacing;
            if (gridWidth <= 0D || double.IsNaN(gridWidth) || double.IsInfinity(gridWidth)) {
                return null;
            }

            return contentWidth.HasValue && contentWidth.Value > 0D
                ? Math.Min(gridWidth, contentWidth.Value)
                : gridWidth;
        }

        private static bool HasNativeTableAuthoredFixedCellWidths(WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    if (cell.Width.GetValueOrDefault() > 0 &&
                        cell.WidthType == W.TableWidthUnitValues.Dxa &&
                        cell.Width.GetValueOrDefault() != NativeOfficeImoScaffoldCellWidthTwips) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ShouldUseNativeAutoFitTableLayout(WordTable table, W.TableProperties? properties, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableLayoutValues? effectiveLayout = properties?.TableLayout?.Type?.Value ?? tableStyleDefaults.Layout;
            if (effectiveLayout == W.TableLayoutValues.Fixed) {
                return false;
            }

            if (effectiveLayout == W.TableLayoutValues.Autofit) {
                return true;
            }

            return !HasNativeTableAuthoredFixedCellWidths(table);
        }

        private static bool IsNativeTableAutoFitToContents(W.TableProperties? properties) =>
            IsNativeTableAutoFitLayout(properties) &&
            properties?.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;

        private static bool IsNativeTableAutoFitLayout(W.TableProperties? properties) {
            if (properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Autofit) {
                return true;
            }

            if (properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Fixed) {
                return false;
            }

            return properties?.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;
        }

        private static double? GetNativeTablePreferredWidth(W.TableWidth? width, double? contentWidth) {
            if (width?.Type?.Value == W.TableWidthUnitValues.Pct) {
                double? percent = GetNativeTablePreferredWidthPercent(width);
                if (!percent.HasValue || !contentWidth.HasValue || contentWidth.Value <= 0D) {
                    return null;
                }

                return contentWidth.Value * percent.Value;
            }

            if (width?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(width.Width?.Value);
        }

        private static double? GetNativeTablePreferredWidthPercent(W.TableWidth width) {
            return GetNativeTableWidthPercent(width.Width?.Value);
        }

        private static double? GetNativeTableCellPreferredWidthPercent(WordTableCell cell) {
            W.TableCellWidth? width = cell._tableCellProperties?.TableCellWidth;
            if (width?.Type?.Value != W.TableWidthUnitValues.Pct) {
                return null;
            }

            return GetNativeTableWidthPercent(width.Width?.Value);
        }

        private static double? GetNativeTableWidthPercent(string? rawWidth) {
            if (string.IsNullOrWhiteSpace(rawWidth)) {
                return null;
            }

            string valueText = rawWidth!.Trim();
            if (valueText.EndsWith("%", StringComparison.Ordinal)) {
                string percentText = valueText.Substring(0, valueText.Length - 1);
                if (!double.TryParse(percentText, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ||
                    percent <= 0D ||
                    double.IsNaN(percent) ||
                    double.IsInfinity(percent)) {
                    return null;
                }

                return percent / 100D;
            }

            if (!int.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) || value <= 0) {
                return null;
            }

            return value / 5000D;
        }

        private static double? GetNativeTableLeftIndent(W.TableIndentation? indentation) {
            if (indentation?.Type?.Value != W.TableWidthUnitValues.Dxa || indentation.Width == null) {
                return null;
            }

            return ConvertNativeTwipsToPoints(indentation.Width.Value);
        }

        private static double? GetNativeTableHorizontalPositionIndent(W.TablePositionProperties? position) {
            if (position?.TablePositionX == null || position.TablePositionXAlignment?.Value != null) {
                return null;
            }

            W.HorizontalAnchorValues? anchor = position.HorizontalAnchor?.Value;
            if (anchor.HasValue &&
                anchor.Value != W.HorizontalAnchorValues.Margin &&
                anchor.Value != W.HorizontalAnchorValues.Text) {
                return null;
            }

            double? indent = ConvertNativeTwipsToPoints(position.TablePositionX.Value);
            return indent.HasValue && indent.Value >= 0D ? indent.Value : null;
        }

        private static double? GetNativeTableCellSpacing(W.TableCellSpacing? spacing) {
            if (spacing?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Width?.Value);
        }

        private static void ApplyNativeTableBorders(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableBorders? directBorders = table._tableProperties?.TableBorders;
            W.TableBorders? tableBorders = directBorders ?? tableStyleDefaults.Borders;
            (PdfCore.PdfColor Color, double Width)? border = directBorders == null
                ? tableStyleDefaults.TableBorder
                : GetNativeUniformTableBorder(directBorders);
            if (border != null) {
                style.BorderColor = border.Value.Color;
                style.BorderWidth = border.Value.Width;
                return;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = CreateNativeTableBorderCellMap(table, tableBorders);
            if (cellBorders == null) {
                return;
            }

            style.BorderColor = null;
            style.BorderWidth = 0D;
            style.CellBorders = cellBorders;
        }

        private static (PdfCore.PdfColor Color, double Width)? GetNativeUniformTableBorder(W.TableBorders? borders) {
            if (borders == null) {
                return null;
            }

            W.BorderType?[] allBorders = {
                borders.TopBorder,
                borders.BottomBorder,
                borders.LeftBorder,
                borders.RightBorder,
                borders.InsideHorizontalBorder,
                borders.InsideVerticalBorder
            };

            if (allBorders.Any(border => border == null || !HasNativeBorder(border.Val?.Value))) {
                return null;
            }

            W.BorderValues style = allBorders[0]!.Val!.Value;
            if (allBorders.Any(border => border!.Val?.Value != style)) {
                return null;
            }

            uint size = allBorders[0]!.Size?.Value ?? 4U;
            if (allBorders.Any(border => (border!.Size?.Value ?? 4U) != size)) {
                return null;
            }

            string? color = NormalizeNativeBorderColor(allBorders[0]!.Color?.Value);
            if (allBorders.Any(border => !string.Equals(color, NormalizeNativeBorderColor(border!.Color?.Value), StringComparison.OrdinalIgnoreCase))) {
                return null;
            }

            return (ParseNativeColor(color) ?? PdfCore.PdfColor.Black, size / 8D);
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateNativeTableBorderCellMap(WordTable table, W.TableBorders? borders) {
            if (!HasNativeTableBorder(borders) || GetNativeUniformTableBorder(borders) != null) {
                return null;
            }

            TableLayout layout = TableLayoutCache.GetLayout(table);
            int rowCount = layout.Rows.Count;
            int columnCount = GetNativeTableColumnCount(layout);
            if (rowCount == 0 || columnCount == 0) {
                return null;
            }

            var cellBorders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                int logicalColumn = GetNativeTableRowStartColumn(layout, rowIndex);
                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    WordTableCell cell = row[cellIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumn += columnSpan;
                        continue;
                    }

                    PdfCore.PdfCellBorder? cellBorder = CreateNativeTableBorderCell(
                        borders!,
                        rowIndex,
                        rowCount,
                        logicalColumn,
                        columnCount,
                        columnSpan,
                        GetNativeCellRowSpan(cell));
                    if (cellBorder != null) {
                        cellBorders[(rowIndex, logicalColumn)] = cellBorder;
                    }

                    logicalColumn += columnSpan;
                }
            }

            return cellBorders.Count == 0 ? null : cellBorders;
        }

        private static void AddNativeTableGridBeforePlaceholders(List<PdfCore.PdfTableCell> cells, int count) {
            for (int i = 0; i < count; i++) {
                cells.Add(PdfCore.PdfTableCell.TextCell(string.Empty));
            }
        }

        private static void AddNativeTableGridAfterPlaceholders(List<PdfCore.PdfTableCell> cells, int count) =>
            AddNativeTableGridBeforePlaceholders(cells, count);

        private static PdfCore.PdfCellBorder? CreateNativeTableBorderCell(W.TableBorders borders, int rowIndex, int rowCount, int columnIndex, int columnCount, int columnSpan, int rowSpan) {
            W.BorderType? top = rowIndex == 0 ? borders.TopBorder : borders.InsideHorizontalBorder;
            W.BorderType? bottom = rowIndex + rowSpan >= rowCount ? borders.BottomBorder : borders.InsideHorizontalBorder;
            W.BorderType? left = columnIndex == 0 ? borders.LeftBorder : borders.InsideVerticalBorder;
            W.BorderType? right = columnIndex + columnSpan >= columnCount ? borders.RightBorder : borders.InsideVerticalBorder;
            bool hasTop = HasNativeBorder(top?.Val?.Value);
            bool hasRight = HasNativeBorder(right?.Val?.Value);
            bool hasBottom = HasNativeBorder(bottom?.Val?.Value);
            bool hasLeft = HasNativeBorder(left?.Val?.Value);
            if (!hasTop && !hasRight && !hasBottom && !hasLeft) {
                return null;
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                Width = 0D,
                TopBorder = CreateNativeCellBorderSide(top),
                RightBorder = CreateNativeCellBorderSide(right),
                BottomBorder = CreateNativeCellBorderSide(bottom),
                LeftBorder = CreateNativeCellBorderSide(left),
                Top = hasTop,
                Right = hasRight,
                Bottom = hasBottom,
                Left = hasLeft
            };
        }

        private static bool HasNativeTableBorder(W.TableBorders? borders) =>
            borders != null &&
            (HasNativeBorder(borders.TopBorder?.Val?.Value) ||
                HasNativeBorder(borders.RightBorder?.Val?.Value) ||
                HasNativeBorder(borders.BottomBorder?.Val?.Value) ||
                HasNativeBorder(borders.LeftBorder?.Val?.Value) ||
                HasNativeBorder(borders.InsideHorizontalBorder?.Val?.Value) ||
                HasNativeBorder(borders.InsideVerticalBorder?.Val?.Value));

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style, bool preserveConfiguredFallbackPadding, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableCellMarginDefault? margins = table._tableProperties?.TableCellMarginDefault;
            if (margins == null) {
                if (tableStyleDefaults.CellPadding != null) {
                    ApplyNativeResolvedTableCellPadding(style, tableStyleDefaults.CellPadding);
                }

                if (!preserveConfiguredFallbackPadding) {
                    style.CellPaddingTop ??= 3D;
                    style.CellPaddingBottom ??= 3D;
                }

                return;
            }

            double? top = ConvertNativeTwipsToPoints(margins.TopMargin?.Width?.Value);
            double? bottom = ConvertNativeTwipsToPoints(margins.BottomMargin?.Width?.Value);
            double? left = margins.TableCellLeftMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellLeftMargin.Width.Value);
            double? right = margins.TableCellRightMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellRightMargin.Width.Value);

            if (top.HasValue) {
                style.CellPaddingTop = top.Value;
            } else if (!preserveConfiguredFallbackPadding) {
                style.CellPaddingTop = 3D;
            }

            if (bottom.HasValue) {
                style.CellPaddingBottom = bottom.Value;
            } else if (!preserveConfiguredFallbackPadding) {
                style.CellPaddingBottom = 3D;
            }

            if (left.HasValue) {
                style.CellPaddingLeft = left.Value;
            }

            if (right.HasValue) {
                style.CellPaddingRight = right.Value;
            }
        }

        private static void ApplyNativeResolvedTableCellPadding(PdfCore.PdfTableStyle style, PdfCore.PdfCellPadding padding) {
            if (padding.Top.HasValue) {
                style.CellPaddingTop = padding.Top.Value;
            }

            if (padding.Bottom.HasValue) {
                style.CellPaddingBottom = padding.Bottom.Value;
            }

            if (padding.Left.HasValue) {
                style.CellPaddingLeft = padding.Left.Value;
            }

            if (padding.Right.HasValue) {
                style.CellPaddingRight = padding.Right.Value;
            }
        }

        private static PdfCore.PdfCellPadding? CreateNativeTableCellPadding(WordTableCell cell) {
            double? top = cell.MarginTopWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginTopWidth.Value) : null;
            double? bottom = cell.MarginBottomWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginBottomWidth.Value) : null;
            double? left = cell.MarginLeftWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginLeftWidth.Value) : null;
            double? right = cell.MarginRightWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginRightWidth.Value) : null;
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right
            };
        }

        private static double? ConvertNativeTwipsToPoints(string? value) {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) || twips < 0) {
                return null;
            }

            return twips / 20D;
        }

        private static double? ConvertNativeTwipsToPoints(int twips) {
            return twips < 0 ? null : twips / 20D;
        }

        private static double ConvertNativeEmusToPoints(long emus) {
            return emus <= 0 ? 0D : emus / 12700D;
        }

        private static void ApplyNativeTableRowOptions(WordTable table, PdfCore.PdfTableStyle style) {
            style.AllowRowBreakAcrossPages = table.AllowRowToBreakAcrossPages;
            List<bool?>? rowBreakPolicies = GetNativeTableRowBreakPolicies(table);
            if (rowBreakPolicies != null) {
                style.RowAllowBreakAcrossPages = rowBreakPolicies;
            }

            List<double?>? rowMinHeights = GetNativeTableRowHeights(table, exact: false);
            if (rowMinHeights != null) {
                double? uniformHeight = GetNativeUniformTableRowHeight(rowMinHeights);
                if (uniformHeight.HasValue) {
                    style.MinRowHeight = uniformHeight.Value;
                } else {
                    style.RowMinHeights = rowMinHeights;
                }
            }

            List<double?>? fixedRowHeights = GetNativeTableRowHeights(table, exact: true);
            if (fixedRowHeights != null) {
                style.FixedRowHeights = fixedRowHeights;
            }
        }

        private static List<bool?>? GetNativeTableRowBreakPolicies(WordTable table) {
            var policies = new List<bool?>(table.Rows.Count);
            bool? firstPolicy = null;
            bool hasMixedPolicies = false;
            foreach (WordTableRow row in table.Rows) {
                bool policy = row.AllowRowToBreakAcrossPages;
                policies.Add(policy);
                if (!firstPolicy.HasValue) {
                    firstPolicy = policy;
                    continue;
                }

                hasMixedPolicies |= firstPolicy.Value != policy;
            }

            return hasMixedPolicies ? policies : null;
        }

        private static List<double?>? GetNativeTableRowHeights(WordTable table, bool exact) {
            var heights = new List<double?>(table.Rows.Count);
            bool hasHeight = false;
            foreach (WordTableRow row in table.Rows) {
                W.TableRowHeight? rowHeight = row._tableRow.TableRowProperties?.Elements<W.TableRowHeight>().FirstOrDefault();
                bool isExact = rowHeight?.HeightType?.Value == W.HeightRuleValues.Exact;
                double? height = row.Height.HasValue && row.Height.Value > 0 && isExact == exact
                    ? ConvertNativeTwipsToPoints(row.Height.Value)
                    : null;
                heights.Add(height);
                hasHeight |= height.HasValue;
            }

            return hasHeight ? heights : null;
        }

        private static double? GetNativeUniformTableRowHeight(IReadOnlyList<double?> rowHeights) {
            double? height = null;
            foreach (double? rowHeight in rowHeights) {
                if (!rowHeight.HasValue) {
                    return null;
                }

                if (!height.HasValue) {
                    height = rowHeight.Value;
                    continue;
                }

                if (System.Math.Abs(height.Value - rowHeight.Value) > 0.001D) {
                    return null;
                }
            }

            return height;
        }

        private static PdfCore.PdfTableStyle? ResolveNativeWordTableStyle(WordTable table, bool preferConfiguredDefaultStyle) {
            string? wordStyle = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(wordStyle)) {
                return null;
            }

            if (preferConfiguredDefaultStyle && IsNativeFallbackTableStyleId(wordStyle)) {
                return null;
            }

            return PdfCore.TableStyles.TryFromWordTableStyle(wordStyle!, out PdfCore.PdfTableStyle? style)
                ? style
                : null;
        }

        private static bool ShouldApplyNativeTableStyleParagraphLineHeight(WordTable table) {
            string? styleId = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            if (!PdfCore.TableStyles.TryGetCanonicalWordStyleName(styleId!, out string? canonicalStyleName)) {
                return true;
            }

            return string.Equals(canonicalStyleName, "TableGrid", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldApplyNativeTableStyleCellPadding(WordTable table) {
            string? styleId = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            if (!PdfCore.TableStyles.TryGetCanonicalWordStyleName(styleId!, out string? canonicalStyleName)) {
                return true;
            }

            return string.Equals(canonicalStyleName, "TableGrid", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(canonicalStyleName, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(canonicalStyleName, "PlainTable1", StringComparison.OrdinalIgnoreCase);
        }

        private static int GetNativeTableVisualHeaderRowCount(WordTable table, int rowCount, int repeatedHeaderRowCount) {
            if (rowCount == 0) {
                return 0;
            }

            int headerRowCount = repeatedHeaderRowCount;
            if (table.ConditionalFormattingFirstRow == true || headerRowCount > 0) {
                headerRowCount = System.Math.Max(headerRowCount, 1);
            }

            return System.Math.Min(headerRowCount, rowCount);
        }

        private static int GetNativeTableRepeatedHeaderRowCount(WordTable table, int rowCount) {
            if (rowCount == 0 || table.Rows.Count == 0) {
                return 0;
            }

            int repeatedHeaderRowCount = 0;
            foreach (WordTableRow row in table.Rows) {
                if (!row.RepeatHeaderRowAtTheTopOfEachPage) {
                    break;
                }

                repeatedHeaderRowCount++;
                if (repeatedHeaderRowCount == rowCount) {
                    break;
                }
            }

            return repeatedHeaderRowCount;
        }

        private static PdfCore.PdfAlign MapNativeTableAlignment(W.TableRowAlignmentValues? alignment) {
            if (alignment == W.TableRowAlignmentValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.TableRowAlignmentValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static W.TableRowAlignmentValues? ResolveNativeTableAlignment(WordTable table, NativeTableStyleDefaults tableStyleDefaults) =>
            ResolveNativeTablePositionAlignment(table._tableProperties?.TablePositionProperties) ??
            table.Alignment ??
            tableStyleDefaults.Alignment;

        private static W.TableRowAlignmentValues? ResolveNativeTablePositionAlignment(W.TablePositionProperties? position) {
            W.HorizontalAlignmentValues? alignment = position?.TablePositionXAlignment?.Value;
            if (alignment == W.HorizontalAlignmentValues.Center) {
                return W.TableRowAlignmentValues.Center;
            }

            if (alignment == W.HorizontalAlignmentValues.Right || alignment == W.HorizontalAlignmentValues.Outside) {
                return W.TableRowAlignmentValues.Right;
            }

            if (alignment == W.HorizontalAlignmentValues.Left || alignment == W.HorizontalAlignmentValues.Inside) {
                return W.TableRowAlignmentValues.Left;
            }

            return null;
        }

        private static List<PdfCore.PdfColumnAlign>? CreateNativeTableHorizontalAlignments(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var alignments = new List<PdfCore.PdfColumnAlign>(columnCount);
            bool hasExplicitAlignment = false;
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PdfCore.PdfColumnAlign? columnAlignment = null;
                bool conflict = false;
                foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                    if (columnIndex < cell.Column || columnIndex >= cell.Column + cell.ColumnSpan) {
                        continue;
                    }

                    PdfCore.PdfColumnAlign alignment = GetNativeCellHorizontalAlignment(cell.Cell);
                    if (columnAlignment == null) {
                        columnAlignment = alignment;
                    } else if (columnAlignment.Value != alignment) {
                        conflict = true;
                        break;
                    }
                }

                PdfCore.PdfColumnAlign resolved = conflict ? PdfCore.PdfColumnAlign.Left : columnAlignment ?? PdfCore.PdfColumnAlign.Left;
                if (resolved != PdfCore.PdfColumnAlign.Left) {
                    hasExplicitAlignment = true;
                }

                alignments.Add(resolved);
            }

            return hasExplicitAlignment ? alignments : null;
        }

        private static PdfCore.PdfColumnAlign GetNativeCellHorizontalAlignment(WordTableCell cell) {
            PdfCore.PdfColumnAlign? alignment = null;
            foreach (WordParagraph paragraph in cell.Paragraphs) {
                string text = GetNativeCellParagraphText(paragraph);
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                PdfCore.PdfColumnAlign paragraphAlignment = ResolveNativeColumnAlign(paragraph);
                if (alignment == null) {
                    alignment = paragraphAlignment;
                } else if (alignment.Value != paragraphAlignment) {
                    return PdfCore.PdfColumnAlign.Left;
                }
            }

            return alignment ?? PdfCore.PdfColumnAlign.Left;
        }

    }
}
