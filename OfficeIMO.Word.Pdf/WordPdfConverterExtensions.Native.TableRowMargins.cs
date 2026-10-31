using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Resolves authored row heights into physical Word cell boxes, including effective vertical cell margins.</summary>
        private static void ApplyNativeUnspacedTableRowMargins(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style) {
            // Spaced cells use their independent border-frame geometry instead.
            if (style.CellSpacing.ToPoints() > 0D || style.BorderFrame?.Border != null) return;
            IReadOnlyList<WordTableRow> sourceRows = table.Rows;
            var minimums = style.RowMinHeights?.ConvertAll(static length => length?.ToPoints()) ?? new List<double?>();
            var fixedHeights = style.FixedRowHeights?.ConvertAll(static length => length?.ToPoints()) ?? new List<double?>();
            while (minimums.Count < sourceRows.Count) minimums.Add(null);
            while (fixedHeights.Count < sourceRows.Count) fixedHeights.Add(null);
            bool hasMinimum = false, hasFixed = false;
            for (int row = 0; row < sourceRows.Count; row++) {
                double top = GetNativeTableRowMargin(layout, style, row, top: true);
                double bottom = GetNativeTableRowMargin(layout, style, row, top: false);
                ApplyNativeRowVerticalMargins(layout, style, row, top, bottom);
                W.TableRowHeight? height = sourceRows[row]._tableRow.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (height?.Val?.Value is not > 0 || height.HeightType?.Value == W.HeightRuleValues.Auto) continue;
                double points = height.Val.Value / 20D;
                if (height.HeightType?.Value == W.HeightRuleValues.Exact) {
                    fixedHeights[row] = points + bottom;
                    hasFixed = true;
                } else {
                    minimums[row] = points + bottom + top +
                        GetNativeTableRowBorderInset(layout, style, row, top: true) +
                        GetNativeTableRowBorderInset(layout, style, row, top: false);
                    hasMinimum = true;
                }
            }
            if (hasFixed) style.FixedRowHeights = (fixedHeights) is { } lengths2322 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths2322, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
            if (hasMinimum) style.RowMinHeights = (minimums) is { } lengths2386 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths2386, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
        }

        private static double GetNativeTableRowBorderInset(TableLayout layout, PdfCore.PdfTableStyle style, int row, bool top) {
            double maximum = 0D;
            int column = GetNativeTableRowStartColumn(layout, row);
            foreach (WordTableCell cell in layout.Rows[row]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    double thickness = style.BorderColor.HasValue ? style.BorderWidth.ToPoints() : 0D;
                    if (style.CellBorders?.TryGetValue((row, column), out PdfCore.PdfCellBorder? border) == true)
                        thickness = GetNativeFrameSide(border, top)?.PaintThickness ?? 0D;
                    maximum = Math.Max(maximum, thickness / 2D);
                }
                column += GetNativeCellColumnSpan(cell);
            }
            return maximum;
        }

        /// <summary>Word neighbors use the row's largest effective vertical margins for their text alignment frame.</summary>
        private static void ApplyNativeRowVerticalMargins(TableLayout layout, PdfCore.PdfTableStyle style, int row, double top, double bottom) {
            int column = layout.GetRowStartColumn(row);
            foreach (WordTableCell cell in layout.Rows[row]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                int span = GetNativeCellColumnSpan(cell);
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    PdfCore.PdfCellPadding? padding = null;
                    style.CellPaddings?.TryGetValue((row, column), out padding);
                    double ownTop = padding?.Top?.ToPoints() ?? style.CellPaddingTop?.ToPoints() ?? style.CellPaddingY.ToPoints();
                    double ownBottom = padding?.Bottom?.ToPoints() ?? style.CellPaddingBottom?.ToPoints() ?? style.CellPaddingY.ToPoints();
                    if (ownTop < top && style.BorderFrame?.Border != null &&
                        style.CellBorders?.TryGetValue((row, column), out PdfCore.PdfCellBorder? border) == true &&
                        border.PaintInsideFrame) {
                        // Text aligns on the shared row margin, but each separate
                        // cell retains its authored top margin inside that frame.
                        PdfCore.PdfCellBorder prepared = border.Clone();
                        prepared.PaintTopInset = top - ownTop;
                        style.CellBorders[(row, column)] = prepared;
                    }
                    if (ownTop < top || ownBottom < bottom) {
                        // Replace the effective PDF value without modifying source cells or caller-owned padding objects.
                        style.CellPaddings ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
                        style.CellPaddings[(row, column)] = new PdfCore.PdfCellPadding {
                            Left = padding?.Left, Right = padding?.Right, Bottom = global::OfficeIMO.OfficeLength.Points(bottom), Top = global::OfficeIMO.OfficeLength.Points(top)
                        };
                    }
                }
                column += span;
            }
        }
    }
}
