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
            double horizontalInset = layout.Rows.Count == 0 ? 0D : GetNativeTableCellBorderInset(style, 0,
                GetNativeTableRowStartColumn(layout, 0), right: false);
            style.BorderFrame = new PdfCore.PdfTableBorderFrame {
                Spacing = style.CellSpacing, Border = border, HorizontalInset = horizontalInset,
                Background = table._tableProperties?.GetFirstChild<W.Shading>() is { } shading
                    ? ParseNativeColor(shading.Fill?.Value) : defaults.CellFill
            };
            ApplyNativeFramedRowHeights(table, layout, style, sourceSpacing, border.TopBorder?.PaintThickness ?? 0D);
        }

        /// <summary>Fixed cells retain a one-point text frame after effective margins and border paint.</summary>
        private static void ApplyNativeFixedTableMinimumWidths(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns) return;
            var minimums = style.ColumnMinWidthPoints == null
                ? new List<double?>() : new List<double?>(style.ColumnMinWidthPoints);
            while (minimums.Count < layout.ColumnWidths.Length) minimums.Add(null);
            var spanningMinimums = new List<(int Column, int Span, double Width)>();
            for (int row = 0; row < layout.Rows.Count; row++) {
                int column = GetNativeTableRowStartColumn(layout, row);
                foreach (WordTableCell cell in layout.Rows[row]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) continue;
                    int span = GetNativeCellColumnSpan(cell);
                    if (!IsNativeVerticalMergeContinuation(cell)) {
                        double width = 1D + GetNativeTableCellHorizontalMargin(style, row, column, right: false) +
                            GetNativeTableCellHorizontalMargin(style, row, column, right: true) +
                            (style.BorderFrame == null ? 0D : 2D * (
                                GetNativeTableCellBorderInset(style, row, column, right: false) +
                                GetNativeTableCellBorderInset(style, row, column, right: true)));
                        double minimum = width - Math.Max(0, span - 1) * style.CellSpacing;
                        if (minimum > 0D) {
                            if (span == 1 && column < minimums.Count)
                                minimums[column] = Math.Max(minimums[column] ?? 0D, minimum);
                            else
                                spanningMinimums.Add((column, span, minimum));
                        }
                    }
                    column += span;
                }
            }
            foreach (var constraint in spanningMinimums)
                ApplyNativeSpanMinimum(layout, minimums, constraint.Column, constraint.Span, constraint.Width);
            style.ColumnMinWidthPoints = minimums;
            if (style.AllowWidthBeyondFrame && style.MaxWidth.HasValue) {
                double minimumWidth = minimums.Sum(width => width ?? 0D) +
                    Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing;
                style.MaxWidth = Math.Max(minimumWidth, style.MaxWidth.Value -
                    (style.BorderFrame == null ? 0D : 2D * style.CellSpacing));
            }
        }

        /// <summary>Fits a merged minimum to the source proportions while retaining single-cell minimums.</summary>
        private static void ApplyNativeSpanMinimum(TableLayout layout, List<double?> minimums,
            int start, int span, double requiredWidth) {
            int end = Math.Min(minimums.Count, start + span);
            double currentWidth = 0D;
            for (int column = start; column < end; column++) currentWidth += minimums[column] ?? 0D;
            if (currentWidth >= requiredWidth) return;
            var weights = new double[end - start];
            double preferredWidth = 0D;
            for (int column = start; column < end; column++)
                preferredWidth += Math.Max(layout.ColumnWidths[column], minimums[column] ?? 0D);
            bool growsBeyondPreferences = requiredWidth > preferredWidth;
            for (int column = start; column < end; column++) {
                double preferred = layout.ColumnWidths[column] > 0F ? layout.ColumnWidths[column] : 1D;
                // Once the complete span outgrows its feasible preferences,
                // minimum frames also participate in proportional growth.
                weights[column - start] = growsBeyondPreferences
                    ? Math.Max(preferred, minimums[column] ?? 0D) + (minimums[column] ?? 0D)
                    : preferred;
            }
            var locked = new bool[end - start];
            double remaining = requiredWidth;
            while (true) {
                double weight = 0D;
                for (int column = start; column < end; column++)
                    if (!locked[column - start]) weight += weights[column - start];
                if (weight <= 0D) return;
                double scale = remaining / weight;
                bool lockedMinimum = false;
                for (int column = start; column < end; column++) {
                    if (locked[column - start]) continue;
                    double candidate = scale * weights[column - start];
                    if (candidate < (minimums[column] ?? 0D)) {
                        locked[column - start] = true;
                        remaining -= minimums[column]!.Value;
                        lockedMinimum = true;
                    }
                }
                if (lockedMinimum) continue;
                for (int column = start; column < end; column++)
                    if (!locked[column - start])
                        minimums[column] = scale * weights[column - start];
                return;
            }
        }

        /// <summary>Translates Word's row constraints into cell-box heights without counting spacing twice.</summary>
        private static void ApplyNativeFramedRowHeights(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style,
            double sourceSpacing, double outerTopThickness) {
            IReadOnlyList<WordTableRow> sourceRows = table.Rows;
            var minimums = style.RowMinHeights == null ? new List<double?>() : new List<double?>(style.RowMinHeights);
            var fixedHeights = style.FixedRowHeights == null ? new List<double?>() : new List<double?>(style.FixedRowHeights);
            while (minimums.Count < sourceRows.Count) minimums.Add(null);
            while (fixedHeights.Count < sourceRows.Count) fixedHeights.Add(null);
            bool hasMinimum = false, hasFixed = false;
            for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) {
                W.TableRowHeight? height = sourceRows[rowIndex]._tableRow.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (height?.Val?.Value is not > 0 || height.HeightType?.Value == W.HeightRuleValues.Auto) continue;
                double top = 0D, bottom = 0D;
                double bottomMargin = GetNativeTableRowMargin(layout, style, rowIndex, top: false);
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

        private static double GetNativeTableRowMargin(TableLayout layout, PdfCore.PdfTableStyle style, int rowIndex, bool top) {
            double maximum = 0D;
            int column = layout.GetRowStartColumn(rowIndex);
            foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                int span = GetNativeCellColumnSpan(cell);
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    PdfCore.PdfCellPadding? padding = null;
                    style.CellPaddings?.TryGetValue((rowIndex, column), out padding);
                    double margin = top
                        ? padding?.Top ?? style.CellPaddingTop ?? style.CellPaddingY
                        : padding?.Bottom ?? style.CellPaddingBottom ?? style.CellPaddingY;
                    maximum = Math.Max(maximum, margin);
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
