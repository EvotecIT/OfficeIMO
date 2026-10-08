using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double NativeAutoFitGridMinimumScale = 0.8D;

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

            double minimumScale = style.AutoFitWidthUsesContentMinimum ? 1D : NativeAutoFitGridMinimumScale;
            if (style.BorderFrame != null)
                tableWidth -= Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing;
            List<double?> derivedMinimums = layout.ColumnWidths
                .Select(width => (double?)(tableWidth * width / gridWidth * minimumScale))
                .ToList();
            if (style.AutoFitWidthUsesContentMinimum) {
                // Only absolute cell preferences reserve the authored grid width.
                // Automatic cells share the remaining width according to content.
                var hasAbsolutePreference = new bool[derivedMinimums.Count];
                foreach (var cell in EnumerateNativeTableCells(layout)) {
                    if (cell.Cell.WidthType != WordTableWidthUnit.Dxa || cell.Cell.Width.GetValueOrDefault() <= 0)
                        continue;
                    for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < hasAbsolutePreference.Length; column++)
                        hasAbsolutePreference[column] = true;
                }
                for (int column = 0; column < derivedMinimums.Count; column++) {
                    if (!hasAbsolutePreference[column])
                        derivedMinimums[column] = null;
                }
            }
            if (style.ColumnMinWidthPoints == null || style.ColumnMinWidthPoints.Count == 0) {
                style.ColumnMinWidthPoints = derivedMinimums;
                return;
            }

            var mergedMinimums = new List<double?>(style.ColumnMinWidthPoints);
            while (mergedMinimums.Count < derivedMinimums.Count) {
                mergedMinimums.Add(null);
            }

            for (int columnIndex = 0; columnIndex < derivedMinimums.Count; columnIndex++) {
                mergedMinimums[columnIndex] ??= derivedMinimums[columnIndex];
            }

            style.ColumnMinWidthPoints = mergedMinimums;
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

        private static void ApplyNativeTableLayoutOptions(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableProperties? properties = table._tableProperties;
            if (ShouldUseNativeAutoFitTableLayout(properties, tableStyleDefaults)) {
                style.AutoFitColumns = true;
            }

            double? cellSpacing = GetNativeTableCellSpacing(properties?.TableCellSpacing) ?? tableStyleDefaults.CellSpacing;
            if (cellSpacing.HasValue) {
                style.CellSpacing = cellSpacing.Value;
            }

            double? preferredWidth = GetNativeTablePreferredWidth(properties?.TableWidth, contentWidth) ??
                GetNativeTablePreferredWidth(tableStyleDefaults.PreferredWidth, contentWidth);
            if (preferredWidth.HasValue && !style.AutoFitColumns) {
                style.MaxWidth = preferredWidth.Value;
                style.PreserveWidth = true;
            } else {
                preferredWidth ??= GetNativeAutoFitGridPreferredWidth(properties, layout, contentWidth, style.CellSpacing);
                if (preferredWidth.HasValue) {
                    // An automatic table's preferred width can grow to fit an
                    // unbreakable cell; it is not a clipping boundary.
                    style.PreferredWidth = preferredWidth.Value;
                    // Positive spacing receives its independent perimeter and
                    // final cell-grid minimums after cell formatting is applied.
                    style.AutoFitWidthUsesContentMinimum = style.AutoFitColumns && style.CellSpacing <= 0D;
                    style.PreserveWidth = true;
                }
            }

            if (style.AutoFitColumns && style.CellSpacing <= 0D && IsNativeTableAutoFitToContents(properties)) {
                // A valid automatic grid can contain only zero/automatic widths.
                // Its physical content minimum still determines the table frame.
                style.AutoFitWidthUsesContentMinimum = true;
                style.PreserveWidth = true;
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

        private static bool ShouldUseNativeAutoFitTableLayout(W.TableProperties? properties, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableLayoutValues? effectiveLayout = properties?.TableLayout?.Type?.Value ?? tableStyleDefaults.Layout;
            if (effectiveLayout == W.TableLayoutValues.Fixed) {
                return false;
            }

            if (effectiveLayout == W.TableLayoutValues.Autofit) {
                return true;
            }

            // Preferred cell widths constrain sizing without changing DOCX's default algorithm.
            return true;
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

            double? points = ConvertNativeTwipsToPoints(width.Width?.Value);
            return points > 0D ? points : null;
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

            return indentation.Width.Value / 20D;
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
            // Word treats these explicit units as zero rather than falling
            // through to a positive spacing value from the table style.
            if (spacing?.Type?.Value == W.TableWidthUnitValues.Auto ||
                spacing?.Type?.Value == W.TableWidthUnitValues.Pct) return 0D;
            if (spacing?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Width?.Value);
        }

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style, bool preserveConfiguredFallbackPadding, NativeTableStyleDefaults tableStyleDefaults) {
            if (!preserveConfiguredFallbackPadding) {
                // Standard Word cell margins apply to each missing side. A
                // partial table declaration must not lose the horizontal inset.
                style.CellPaddingTop ??= 0D;
                style.CellPaddingBottom ??= 0D;
                style.CellPaddingLeft ??= 5.4D;
                style.CellPaddingRight ??= 5.4D;
            }
            if (tableStyleDefaults.CellPadding != null) {
                ApplyNativeResolvedTableCellPadding(style, tableStyleDefaults.CellPadding);
            }
            W.TableCellMarginDefault? margins = table._tableProperties?.TableCellMarginDefault;
            if (margins == null) {
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
            }

            if (bottom.HasValue) {
                style.CellPaddingBottom = bottom.Value;
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
            table.Alignment.ToOpenXml() ??
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
