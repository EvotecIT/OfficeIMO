using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeTableConditionalPaddings(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, PdfCore.PdfTableStyle style) {
            if (layout.Rows.Count == 0) {
                return;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> cellPaddings = style.CellPaddings == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(style.CellPaddings);
            bool changed = false;
            if (table.ConditionalFormattingFirstRow == true && style.HeaderRowCount > 0) {
                int headerRows = System.Math.Min(style.HeaderRowCount, layout.Rows.Count);
                for (int rowIndex = 0; rowIndex < headerRows; rowIndex++) {
                    changed |= ApplyNativeTableConditionalRowPadding(layout, rowIndex, tableStyleDefaults.FirstRowStyle, cellPaddings);
                }
            }

            if (table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount) {
                changed |= ApplyNativeTableConditionalRowPadding(layout, layout.Rows.Count - 1, tableStyleDefaults.LastRowStyle, cellPaddings);
            }

            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount > 0) {
                int footerStartRowIndex = table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount
                    ? layout.Rows.Count - 1
                    : layout.Rows.Count;
                changed |= ApplyNativeTableConditionalBandPaddings(table, layout, tableStyleDefaults, style.HeaderRowCount, footerStartRowIndex, cellPaddings);
                changed |= ApplyNativeTableConditionalColumnPaddings(table, layout, tableStyleDefaults, columnCount, cellPaddings);
            }

            if (changed) {
                style.CellPaddings = cellPaddings;
            }
        }

        private static bool ApplyNativeTableConditionalRowPadding(TableLayout layout, int rowIndex, NativeTableConditionalStyleDefaults conditionalStyle, Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> cellPaddings) {
            if (conditionalStyle.CellPadding == null) {
                return false;
            }

            bool changed = false;
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

                changed |= ApplyNativeTableConditionalCellPadding(cellPaddings, (rowIndex, logicalColumnIndex), conditionalStyle.CellPadding);
                logicalColumnIndex += columnSpan;
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalBandPaddings(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int headerRowCount, int footerStartRowIndex, Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> cellPaddings) {
            PdfCore.PdfCellPadding? horizontalBandPadding = table.ConditionalFormattingNoHorizontalBand != true
                ? tableStyleDefaults.Band1HorizontalStyle.CellPadding
                : null;
            PdfCore.PdfCellPadding? verticalBandPadding = table.ConditionalFormattingNoVerticalBand != true
                ? tableStyleDefaults.Band1VerticalStyle.CellPadding
                : null;
            if (horizontalBandPadding == null && verticalBandPadding == null) {
                return false;
            }

            bool changed = false;
            for (int rowIndex = headerRowCount; rowIndex < footerStartRowIndex; rowIndex++) {
                int bodyRowIndex = rowIndex - headerRowCount;
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
                    if (bodyRowIndex % 2 == 1) {
                        changed |= ApplyNativeTableConditionalCellPadding(cellPaddings, key, horizontalBandPadding);
                    }

                    if (logicalColumnIndex % 2 == 1) {
                        changed |= ApplyNativeTableConditionalCellPadding(cellPaddings, key, verticalBandPadding);
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalColumnPaddings(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int columnCount, Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> cellPaddings) {
            PdfCore.PdfCellPadding? firstColumnPadding = table.ConditionalFormattingFirstColumn == true
                ? tableStyleDefaults.FirstColumnStyle.CellPadding
                : null;
            PdfCore.PdfCellPadding? lastColumnPadding = table.ConditionalFormattingLastColumn == true
                ? tableStyleDefaults.LastColumnStyle.CellPadding
                : null;
            if (firstColumnPadding == null && lastColumnPadding == null) {
                return false;
            }

            bool changed = false;
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
                    if (logicalColumnIndex == 0) {
                        changed |= ApplyNativeTableConditionalCellPadding(cellPaddings, key, firstColumnPadding);
                    }

                    if (logicalColumnIndex + columnSpan >= columnCount) {
                        changed |= ApplyNativeTableConditionalCellPadding(cellPaddings, key, lastColumnPadding);
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalCellPadding(Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> cellPaddings, (int Row, int Column) key, PdfCore.PdfCellPadding? padding) {
            if (padding == null) {
                return false;
            }

            cellPaddings[key] = MergeNativeCellPadding(cellPaddings.TryGetValue(key, out PdfCore.PdfCellPadding? existing) ? existing : null, padding)!;
            return true;
        }

        private static PdfCore.PdfCellPadding? MergeNativeCellPadding(PdfCore.PdfCellPadding? basePadding, PdfCore.PdfCellPadding? overlayPadding) {
            if (basePadding == null) {
                return overlayPadding?.Clone();
            }

            if (overlayPadding == null) {
                return basePadding.Clone();
            }

            return new PdfCore.PdfCellPadding {
                Top = overlayPadding.Top ?? basePadding.Top,
                Bottom = overlayPadding.Bottom ?? basePadding.Bottom,
                Left = overlayPadding.Left ?? basePadding.Left,
                Right = overlayPadding.Right ?? basePadding.Right
            };
        }
    }
}
