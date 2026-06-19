using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeTableConditionalRowBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, PdfCore.PdfTableStyle style) {
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
            int logicalColumnIndex = 0;
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
