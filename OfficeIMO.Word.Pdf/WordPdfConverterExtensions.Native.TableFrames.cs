using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeTableBorderFrame(WordTable table, TableLayout layout,
            PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
            if (style.CellSpacing <= 0D || style.Position != null || !style.ConsumesVerticalFlow) return;
            W.TableBorders? authored = table._tableProperties?.TableBorders;
            W.TableBorders? borders = authored?.HasChildren == true
                ? MergeNativeTableBorders(defaults.Borders, authored) : defaults.Borders;
            var border = new PdfCore.PdfCellBorder {
                Color = null, Width = 0D,
                Top = HasNativeBorder(borders?.TopBorder?.Val?.Value),
                Right = HasNativeBorder(borders?.RightBorder?.Val?.Value),
                Bottom = HasNativeBorder(borders?.BottomBorder?.Val?.Value),
                Left = HasNativeBorder(borders?.LeftBorder?.Val?.Value),
                TopBorder = CreateNativeCellBorderSide(borders?.TopBorder),
                RightBorder = CreateNativeCellBorderSide(borders?.RightBorder),
                BottomBorder = CreateNativeCellBorderSide(borders?.BottomBorder),
                LeftBorder = CreateNativeCellBorderSide(borders?.LeftBorder)
            };
            MaterializeNativeTableBorderGrid(style, layout);
            if (style.CellBorders != null)
                foreach (PdfCore.PdfCellBorder cellBorder in style.CellBorders.Values) cellBorder.PaintInsideFrame = true;
            double sourceSpacing = style.CellSpacing;
            style.CellSpacing = 2D * sourceSpacing;
            if (style.PreferredWidth.HasValue)
                style.PreferredWidth += Math.Max(0, layout.ColumnWidths.Length - 1) * sourceSpacing;
            style.AutoFitWidthUsesContentMinimum = style.AutoFitColumns;
            style.CellVerticalPaddingFromBorderInterior = true;
            style.PreservePartialCellLines = true;
            // The independent perimeter supplies the source continuation inset.
            style.PageContinuationSpacingBefore = 0D;
            double horizontalInset = 0D;
            if (style.CellBorders != null) {
                foreach (var entry in style.CellBorders.Where(entry => entry.Key.Row == 0)) {
                    horizontalInset = Math.Max(horizontalInset,
                        Math.Max(GetNativeFrameSide(entry.Value, top: null, right: false)?.PaintThickness ?? 0D,
                            GetNativeFrameSide(entry.Value, top: null, right: true)?.PaintThickness ?? 0D) / 2D);
                }
            }
            style.BorderFrame = new PdfCore.PdfTableBorderFrame {
                Spacing = style.CellSpacing, Border = border, HorizontalInset = horizontalInset,
                Background = table._tableProperties?.GetFirstChild<W.Shading>() is { } shading
                    ? ParseNativeColor(shading.Fill?.Value) : defaults.CellFill
            };
            ApplyNativeFramedRowHeights(table, layout, style, sourceSpacing, border.TopBorder?.PaintThickness ?? 0D);
        }

        /// <summary>Translates Word's row constraints into cell-box heights without counting spacing twice.</summary>
        private static void ApplyNativeFramedRowHeights(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style,
            double sourceSpacing, double outerTopThickness) {
            var minimums = style.RowMinHeights == null ? new List<double?>() : new List<double?>(style.RowMinHeights);
            var fixedHeights = style.FixedRowHeights == null ? new List<double?>() : new List<double?>(style.FixedRowHeights);
            while (minimums.Count < table.Rows.Count) minimums.Add(null);
            while (fixedHeights.Count < table.Rows.Count) fixedHeights.Add(null);
            bool hasMinimum = false, hasFixed = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                W.TableRowHeight? height = table.Rows[rowIndex]._tableRow.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (height?.Val?.Value is not > 0 || height.HeightType?.Value == W.HeightRuleValues.Auto) continue;
                double top = 0D, bottom = 0D;
                double bottomMargin = GetNativeFramedRowBottomMargin(layout, style, rowIndex);
                if (style.CellBorders != null) {
                    foreach (var entry in style.CellBorders.Where(entry => entry.Key.Row == rowIndex)) {
                        top = Math.Max(top, GetNativeFrameSide(entry.Value, top: true)?.PaintThickness ?? 0D);
                        bottom = Math.Max(bottom, GetNativeFrameSide(entry.Value, top: false)?.PaintThickness ?? 0D);
                    }
                }
                double points = height.Val.Value / 20D;
                if (height.HeightType?.Value == W.HeightRuleValues.Exact) {
                    // The first row includes leading spacing; subsequent rows
                    // share the half-spacing and the adjoining border paint.
                    double spacing = rowIndex == 0 ? 2D * sourceSpacing : sourceSpacing;
                    double borderHeight = rowIndex == 0 ? -Math.Max(0D, outerTopThickness - top) : (top + bottom) / 2D;
                    fixedHeights[rowIndex] = Math.Max(0D, points - spacing + borderHeight + bottomMargin);
                    hasFixed = true;
                } else {
                    minimums[rowIndex] = points + top + bottom;
                    hasMinimum = true;
                }
            }
            if (hasFixed) style.FixedRowHeights = fixedHeights;
            if (hasMinimum) { style.MinRowHeight = 0D; style.RowMinHeights = minimums; }
        }

        private static double GetNativeFramedRowBottomMargin(TableLayout layout, PdfCore.PdfTableStyle style, int rowIndex) {
            double maximum = 0D;
            int column = layout.GetRowStartColumn(rowIndex);
            foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                int span = GetNativeCellColumnSpan(cell);
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    PdfCore.PdfCellPadding? padding = null;
                    style.CellPaddings?.TryGetValue((rowIndex, column), out padding);
                    maximum = Math.Max(maximum, padding?.Bottom ?? style.CellPaddingBottom ?? style.CellPaddingY);
                }
                column += span;
            }
            return maximum;
        }

        private static PdfCore.PdfCellBorderSide? GetNativeFrameSide(PdfCore.PdfCellBorder border,
            bool? top, bool right = false) {
            bool enabled = top.HasValue ? top.Value ? border.Top : border.Bottom : right ? border.Right : border.Left;
            if (!enabled) return null;
            PdfCore.PdfCellBorderSide? side = top.HasValue ? top.Value ? border.TopBorder : border.BottomBorder
                : right ? border.RightBorder : border.LeftBorder;
            return side ?? new PdfCore.PdfCellBorderSide {
                Color = border.Color, Width = border.Width, DashStyle = border.DashStyle, LineStyle = border.LineStyle
            };
        }
    }
}
