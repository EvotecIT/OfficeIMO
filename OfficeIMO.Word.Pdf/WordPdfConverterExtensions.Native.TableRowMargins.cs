using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Resolves authored row heights into physical Word cell boxes, including effective vertical cell margins.</summary>
        private static void ApplyNativeUnspacedTableRowMargins(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style) {
            // Spaced cells use their independent border-frame geometry instead.
            if (style.CellSpacing > 0D) return;
            IReadOnlyList<WordTableRow> sourceRows = table.Rows;
            var minimums = style.RowMinHeights == null ? new List<double?>() : new List<double?>(style.RowMinHeights);
            var fixedHeights = style.FixedRowHeights == null ? new List<double?>() : new List<double?>(style.FixedRowHeights);
            while (minimums.Count < sourceRows.Count) minimums.Add(null);
            while (fixedHeights.Count < sourceRows.Count) fixedHeights.Add(null);
            bool hasMinimum = false, hasFixed = false;
            for (int row = 0; row < sourceRows.Count; row++) {
                double top = GetNativeTableRowMargin(layout, style, row, top: true);
                ApplyNativeRowTopMargin(layout, style, row, top);
                W.TableRowHeight? height = sourceRows[row]._tableRow.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (height?.Val?.Value is not > 0 || height.HeightType?.Value == W.HeightRuleValues.Auto) continue;
                double points = height.Val.Value / 20D;
                double bottom = GetNativeTableRowMargin(layout, style, row, top: false);
                if (height.HeightType?.Value == W.HeightRuleValues.Exact) {
                    fixedHeights[row] = points + bottom;
                    hasFixed = true;
                } else {
                    minimums[row] = points + bottom + top;
                    hasMinimum = true;
                }
            }
            if (hasFixed) style.FixedRowHeights = fixedHeights;
            if (hasMinimum) style.RowMinHeights = minimums;
        }

        /// <summary>Word top-aligned neighbors share the row's largest top margin.</summary>
        private static void ApplyNativeRowTopMargin(TableLayout layout, PdfCore.PdfTableStyle style, int row, double top) {
            int column = layout.GetRowStartColumn(row);
            foreach (WordTableCell cell in layout.Rows[row]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                int span = GetNativeCellColumnSpan(cell);
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    PdfCore.PdfCellPadding? padding = null;
                    style.CellPaddings?.TryGetValue((row, column), out padding);
                    double ownTop = padding?.Top ?? style.CellPaddingTop ?? style.CellPaddingY;
                    if (ownTop < top) {
                        // Replace the effective PDF value without modifying source cells or caller-owned padding objects.
                        style.CellPaddings ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
                        style.CellPaddings[(row, column)] = new PdfCore.PdfCellPadding {
                            Left = padding?.Left, Right = padding?.Right, Bottom = padding?.Bottom, Top = top
                        };
                    }
                }
                column += span;
            }
        }
    }
}
