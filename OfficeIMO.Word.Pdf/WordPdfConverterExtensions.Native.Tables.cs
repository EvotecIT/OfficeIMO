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
        private static void RenderNativeTable(INativePdfFlow pdf, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, WordToPdfOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            RecordNativeBodyTableDiagnostics(table, options, "body table");

            TableLayout layout = TableLayoutCache.GetLayout(table);
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
            var rows = new List<PdfCore.PdfTableCell[]>();
            var cellFills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            var directCellBorders = new Dictionary<(int Row, int Column), WordTableCellBorder>();
            var cellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
            var cellAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            var cellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
            NativeTableColumnAlignments tableAlignments = CreateNativeTableColumnAlignments(layout);
            List<PdfCore.PdfColumnAlign>? horizontalAlignments = tableAlignments.Horizontal;
            List<PdfCore.PdfCellVerticalAlign>? verticalAlignments = tableAlignments.Vertical;
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
                    NativeCellText cellText = CreateNativeCellText(
                        cell,
                        footnoteNumbersById,
                        nativeDefaults,
                        cellStyleDefaults,
                        nativeFontMap,
                        getMarker,
                        ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
                    NativeTableCellEmbeddedContent embeddedContent = CreateNativeTableCellEmbeddedContent(cell, options);
                    (string? LinkUri, string? LinkContents) link = GetNativeCellLink(cell);
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellText.Runs,
                        cellText.Paragraphs,
                        columnSpan,
                        link.LinkUri,
                        link.LinkContents,
                        rowSpan,
                        embeddedContent.CheckBoxes.Count == 0 ? null : embeddedContent.CheckBoxes,
                        embeddedContent.FormFields.Count == 0 ? null : embeddedContent.FormFields,
                        embeddedContent.Images.Count == 0 ? null : embeddedContent.Images,
                        noWrap: !cell.WrapText));

                    PdfCore.PdfColor? fill =
                        ParseNativeColor(cell.ShadingFillColorHex) ??
                        cellStyleDefaults.CellFill;
                    if (fill.HasValue) {
                        cellFills[(rowIndex, logicalColumnIndex)] = fill.Value;
                    }

                    if (HasNativeDirectCellBorder(cell.Borders)) {
                        directCellBorders[(rowIndex, logicalColumnIndex)] = cell.Borders;
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
            if (table._tableProperties?.TablePositionProperties is { } tablePosition) {
                style.ConsumesVerticalFlow = false;
                if (!pdf.SupportsPositionedTables && options != null) {
                    AddNativeExportWarning(options,
                        "NativePositionedTableWrapApproximation",
                        "table",
                        "Positioned table wrapping in a multi-column section is approximated.");
                }
                if (pdf.SupportsPositionedTables) {
                    style.Position = CreateNativeTablePosition(tablePosition,
                        table._tableProperties?.GetFirstChild<W.TableOverlap>()?.Val?.Value != W.TableOverlapValues.Never);
                    // tblpX is a placement coordinate, not a width reservation.
                    style.LeftIndent = 0;
                }
            }
            if (cellFills.Count > 0) {
                if (style.CellFills == null) {
                    style.CellFills = cellFills;
                } else {
                    foreach (var cellFill in cellFills) {
                        style.CellFills[cellFill.Key] = cellFill.Value;
                    }
                }
            }

            ApplyNativeDirectCellBorders(style, layout, directCellBorders);
            if (style.CellBorders != null && style.CellSpacing <= 0D) {
                ReconcileNativeHiddenSharedBorders(table, layout, tableStyleDefaults, style.HeaderRowCount, style.CellBorders, directCellBorders);
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

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, WordToPdfOptions? options) =>
            CreateNativeTableStyle(table, rowCount, options, null);

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, WordToPdfOptions? options, double? contentWidth) =>
            CreateNativeTableStyle(table, rowCount, options, contentWidth, GetNativeDocumentDefaults(table.Document));

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, WordToPdfOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults) {
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
            WordToPdfOptions? options,
            double? contentWidth,
            NativeDocumentDefaults nativeDefaults,
            NativeTableStyleDefaults tableStyleDefaults,
            TableLayout layout,
            NativeFontMap? nativeFontMap = null) {
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            PdfCore.PdfTableStyle? wordStyle = ResolveNativeWordTableStyle(table, hasExplicitDefaultTableStyle);
            bool usesConfiguredDefaultStyle = wordStyle == null && hasExplicitDefaultTableStyle;
            PdfCore.PdfTableStyle style = wordStyle ?? CreateNativeDefaultTableStyle(options);
            style.ClipTextToCellBounds = true;
            if (!usesConfiguredDefaultStyle) {
                // Word tables do not infer a two-row first/last-fragment group from
                // the shared renderer's presentation defaults. Source paragraph and
                // row rules govern pagination unless a PDF table policy is configured.
                style.MinimumBodyRowsOnFirstPage = 0;
                style.MinimumBodyRowsOnLastPage = 0;
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
            if (!usesConfiguredDefaultStyle && repeatedHeaderRowCount > 0) {
                // A repeated source header starts with body content. Removing
                // presentation row groups must not leave a header by itself.
                style.MinimumBodyRowsOnFirstPage = 1;
            }
            if (repeatedHeaderRowCount > 0) {
                style.PageContinuationSpacingBefore = Math.Max(style.PageContinuationSpacingBefore, NativeTablePageContinuationSpacingBefore);
            }

            if (options?.DefaultTableBorders == true && style.BorderColor == null) {
                style.BorderColor = PdfCore.PdfColor.LightGray;
                if (style.BorderWidth <= 0D) {
                    style.BorderWidth = 0.5D;
                }
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
            if (HasNativeConditionalHiddenBorders(table, tableStyleDefaults)) {
                MaterializeNativeTableBorderGrid(style, layout);
            }
            ApplyNativeTableConditionalBorders(table, layout, tableStyleDefaults, style);
            ApplyNativeTableConditionalPaddings(table, layout, tableStyleDefaults, style);
            ApplyNativeTableLayoutOptions(table, layout, style, contentWidth, tableStyleDefaults);
            style.CellVerticalPaddingFromBorderInterior = !usesConfiguredDefaultStyle && style.CellSpacing <= 0D;
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
            return tableStyleDefaults.LineSpacing.Resolve(fontSize,
                ResolveNativeWordSingleLineHeight(nativeFontMap, tableStyleDefaults.RunStyle.FontFamily, documentFontFamily))
                ?? tableStyleDefaults.ParagraphLineHeight;
        }

        private static PdfCore.PdfTableStyle CreateNativeDefaultTableStyle(WordToPdfOptions? options) {
            PdfCore.PdfTableStyle? configuredStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true
                ? options.PdfOptions.DefaultTableStyle
                : null;
            if (configuredStyle != null) {
                return configuredStyle.Clone();
            }

            return new PdfCore.PdfTableStyle {
                BorderColor = null,
                BorderWidth = 0D,
                HeaderFill = null,
                FooterFill = null,
                HeaderBold = false,
                FooterBold = false,
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
            if (!conditionalStyle.CellFill.HasValue &&
                conditionalStyle.ParagraphPagination == default &&
                !conditionalStyle.TextColor.HasValue &&
                !conditionalStyle.FontSize.HasValue &&
                !conditionalStyle.ComplexScript.FontSize.HasValue &&
                !conditionalStyle.ComplexScript.Enabled.HasValue &&
                string.IsNullOrWhiteSpace(conditionalStyle.FontFamily) &&
                !conditionalStyle.Bold.HasValue &&
                !conditionalStyle.Italic.HasValue &&
                !conditionalStyle.UnderlineStyle.HasValue &&
                !conditionalStyle.StrikeStyle.HasValue &&
                !conditionalStyle.AllCaps.HasValue &&
                !conditionalStyle.Baseline.HasValue &&
                !conditionalStyle.Highlight.HasValue &&
                !conditionalStyle.CellVerticalAlignment.HasValue &&
                !conditionalStyle.ParagraphLineHeight.HasValue &&
                !conditionalStyle.ParagraphLineSpacingPoints.HasValue &&
                !conditionalStyle.ParagraphLineSpacingRule.HasValue &&
                !conditionalStyle.LineSpacing.Value.HasValue && !conditionalStyle.LineSpacing.Rule.HasValue &&
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
                CellFill = conditionalStyle.CellFill ?? tableStyleDefaults.CellFill,
                CellVerticalAlignment = conditionalStyle.CellVerticalAlignment ?? tableStyleDefaults.CellVerticalAlignment,
                ParagraphLineHeight = conditionalStyle.ParagraphLineHeight ?? tableStyleDefaults.ParagraphLineHeight,
                ParagraphLineSpacingPoints = conditionalStyle.ParagraphLineSpacingPoints ?? tableStyleDefaults.ParagraphLineSpacingPoints,
                ParagraphLineSpacingRule = conditionalStyle.ParagraphLineSpacingRule ?? tableStyleDefaults.ParagraphLineSpacingRule,
                LineSpacing = conditionalStyle.LineSpacing.Inherit(tableStyleDefaults.LineSpacing),
                ParagraphPagination = conditionalStyle.ParagraphPagination.Inherit(tableStyleDefaults.ParagraphPagination),
                ParagraphSpacingBefore = conditionalStyle.ParagraphSpacingBefore ?? tableStyleDefaults.ParagraphSpacingBefore,
                ParagraphSpacingAfter = conditionalStyle.ParagraphSpacingAfter ?? tableStyleDefaults.ParagraphSpacingAfter,
                ParagraphAlignment = conditionalStyle.ParagraphAlignment ?? tableStyleDefaults.ParagraphAlignment,
                ParagraphLeftIndent = conditionalStyle.ParagraphLeftIndent ?? tableStyleDefaults.ParagraphLeftIndent,
                ParagraphRightIndent = conditionalStyle.ParagraphRightIndent ?? tableStyleDefaults.ParagraphRightIndent,
                ParagraphFirstLineIndent = conditionalStyle.ParagraphFirstLineIndent ?? tableStyleDefaults.ParagraphFirstLineIndent,
                RunStyle = runStyle with {
                    FontSize = conditionalStyle.FontSize ?? runStyle.FontSize,
                    ComplexScript = runStyle.ComplexScript.Merge(conditionalStyle.ComplexScript),
                    FontFamily = conditionalStyle.FontFamily ?? runStyle.FontFamily,
                    Bold = conditionalStyle.Bold ?? runStyle.Bold,
                    Italic = conditionalStyle.Italic ?? runStyle.Italic,
                    UnderlineStyle = conditionalStyle.UnderlineStyle ?? runStyle.UnderlineStyle,
                    StrikeStyle = conditionalStyle.StrikeStyle ?? runStyle.StrikeStyle,
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

        private static void ApplyNativeTableBorders(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableBorders? directBorders = table._tableProperties?.TableBorders;
            if (directBorders?.HasChildren != true) directBorders = null;
            W.TableBorders? tableBorders = directBorders == null
                ? tableStyleDefaults.Borders
                : MergeNativeTableBorders(tableStyleDefaults.Borders, directBorders);
            (PdfCore.PdfColor Color, double Width)? border = directBorders == null
                ? tableStyleDefaults.TableBorder
                : GetNativeUniformTableBorder(tableBorders);
            if (border != null) {
                style.BorderColor = border.Value.Color;
                style.BorderWidth = border.Value.Width;
                return;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = CreateNativeTableBorderCellMap(table, tableBorders);
            if (cellBorders == null) {
                if (directBorders != null) {
                    style.BorderColor = null;
                    style.BorderWidth = 0D;
                    style.CellBorders = null;
                }
                return;
            }

            style.BorderColor = null;
            style.BorderWidth = 0D;
            style.CellBorders = cellBorders;
        }

        private static W.TableBorders MergeNativeTableBorders(W.TableBorders? inherited, W.TableBorders direct) {
            var merged = new W.TableBorders();
            AppendBorder(direct.TopBorder ?? inherited?.TopBorder);
            AppendBorder(direct.LeftBorder ?? inherited?.LeftBorder);
            AppendBorder(direct.BottomBorder ?? inherited?.BottomBorder);
            AppendBorder(direct.RightBorder ?? inherited?.RightBorder);
            AppendBorder(direct.InsideHorizontalBorder ?? inherited?.InsideHorizontalBorder);
            AppendBorder(direct.InsideVerticalBorder ?? inherited?.InsideVerticalBorder);
            return merged;

            void AppendBorder(W.BorderType? source) {
                if (source != null) {
                    merged.Append(source.CloneNode(true));
                }
            }
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
            // The uniform grid stores only color and width. Patterned and paired
            // strokes need the existing per-cell border representation.
            if (ToNativeBorderDashStyle(style) != OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid ||
                ToNativeBorderLineStyle(style) != PdfCore.PdfCellBorderLineStyle.Standard) {
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

    }
}
