using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeTableConditionalBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, PdfCore.PdfTableStyle style) {
            if (layout.Rows.Count == 0) {
                return;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = style.CellBorders == null
                ? null
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>(style.CellBorders);
            cellBorders ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            bool changed = false;

            if (table.ConditionalFormattingFirstRow == true) {
                changed |= ApplyNativeTableConditionalRowBorder(layout, 0, tableStyleDefaults.FirstRowStyle, cellBorders);
            }

            if (table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount) {
                changed |= ApplyNativeTableConditionalRowBorder(layout, layout.Rows.Count - 1, tableStyleDefaults.LastRowStyle, cellBorders);
            }

            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount > 0) {
                int footerStartRowIndex = table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount
                    ? layout.Rows.Count - 1
                    : layout.Rows.Count;
                changed |= ApplyNativeTableConditionalBandBorders(table, layout, tableStyleDefaults, style.HeaderRowCount, footerStartRowIndex, cellBorders);
                changed |= ApplyNativeTableConditionalColumnBorders(table, layout, tableStyleDefaults, columnCount, cellBorders);
            }

            if (changed) {
                style.CellBorders = cellBorders;
            }
        }

        private static bool ApplyNativeTableConditionalRowBorder(TableLayout layout, int rowIndex, NativeTableConditionalStyleDefaults conditionalStyle, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            PdfCore.PdfCellBorder? border = CreateNativeTableConditionalCellBorder(conditionalStyle.CellBorders);
            if (border == null) {
                return false;
            }

            IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
            bool changed = false;
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
                cellBorders[key] = MergeNativeCellBorder(cellBorders.TryGetValue(key, out PdfCore.PdfCellBorder? existing) ? existing : null, border);
                changed = true;
                logicalColumnIndex += columnSpan;
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalBandBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int headerRowCount, int footerStartRowIndex, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            PdfCore.PdfCellBorder? horizontalBandBorder = table.ConditionalFormattingNoHorizontalBand != true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.Band1HorizontalStyle.CellBorders)
                : null;
            PdfCore.PdfCellBorder? verticalBandBorder = table.ConditionalFormattingNoVerticalBand != true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.Band1VerticalStyle.CellBorders)
                : null;
            if (horizontalBandBorder == null && verticalBandBorder == null) {
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
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, horizontalBandBorder);
                    }

                    if (logicalColumnIndex % 2 == 1) {
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, verticalBandBorder);
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalColumnBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int columnCount, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            PdfCore.PdfCellBorder? firstColumnBorder = table.ConditionalFormattingFirstColumn == true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.FirstColumnStyle.CellBorders)
                : null;
            PdfCore.PdfCellBorder? lastColumnBorder = table.ConditionalFormattingLastColumn == true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.LastColumnStyle.CellBorders)
                : null;
            if (firstColumnBorder == null && lastColumnBorder == null) {
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
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, firstColumnBorder);
                    }

                    if (logicalColumnIndex + columnSpan >= columnCount) {
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, lastColumnBorder);
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalCellBorder(Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders, (int Row, int Column) key, PdfCore.PdfCellBorder? border) {
            if (border == null) {
                return false;
            }

            cellBorders[key] = MergeNativeCellBorder(cellBorders.TryGetValue(key, out PdfCore.PdfCellBorder? existing) ? existing : null, border);
            return true;
        }

        private static PdfCore.PdfCellBorder? CreateNativeTableConditionalCellBorder(W.TableCellBorders? borders) {
            if (borders == null) {
                return null;
            }

            W.BorderType? top = borders.GetFirstChild<W.TopBorder>();
            W.BorderType? right = borders.GetFirstChild<W.RightBorder>();
            right ??= borders.GetFirstChild<W.EndBorder>();
            W.BorderType? bottom = borders.GetFirstChild<W.BottomBorder>();
            W.BorderType? left = borders.GetFirstChild<W.LeftBorder>();
            left ??= borders.GetFirstChild<W.StartBorder>();
            bool hasTop = HasNativeConditionalBorderOverride(top);
            bool hasRight = HasNativeConditionalBorderOverride(right);
            bool hasBottom = HasNativeConditionalBorderOverride(bottom);
            bool hasLeft = HasNativeConditionalBorderOverride(left);
            if (!hasTop && !hasRight && !hasBottom && !hasLeft) {
                return null;
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                Width = 0D,
                TopBorder = CreateNativeConditionalBorderSide(top),
                RightBorder = CreateNativeConditionalBorderSide(right),
                BottomBorder = CreateNativeConditionalBorderSide(bottom),
                LeftBorder = CreateNativeConditionalBorderSide(left),
                Top = hasTop,
                Right = hasRight,
                Bottom = hasBottom,
                Left = hasLeft
            };
        }

        private static bool HasNativeConditionalBorderOverride(W.BorderType? border) =>
            HasNativeBorder(border?.Val?.Value) ||
            border?.Val?.Value == W.BorderValues.Nil ||
            border?.Val?.Value == W.BorderValues.None;

        private static PdfCore.PdfCellBorderSide? CreateNativeConditionalBorderSide(W.BorderType? border) {
            if (border?.Val?.Value == W.BorderValues.Nil || border?.Val?.Value == W.BorderValues.None) {
                return new PdfCore.PdfCellBorderSide { Color = null, Width = 0D };
            }

            return CreateNativeCellBorderSide(border);
        }

        private static PdfCore.PdfCellBorder MergeNativeCellBorder(PdfCore.PdfCellBorder? existing, PdfCore.PdfCellBorder overlay) {
            PdfCore.PdfCellBorder result = existing?.Clone() ?? new PdfCore.PdfCellBorder {
                Color = null,
                Width = 0D,
                Top = false,
                Right = false,
                Bottom = false,
                Left = false,
                DiagonalUp = false,
                DiagonalDown = false
            };

            if (overlay.Top) {
                result.Top = true;
                result.TopBorder = overlay.TopBorder;
            }

            if (overlay.Right) {
                result.Right = true;
                result.RightBorder = overlay.RightBorder;
            }

            if (overlay.Bottom) {
                result.Bottom = true;
                result.BottomBorder = overlay.BottomBorder;
            }

            if (overlay.Left) {
                result.Left = true;
                result.LeftBorder = overlay.LeftBorder;
            }

            return result;
        }
    }
}
