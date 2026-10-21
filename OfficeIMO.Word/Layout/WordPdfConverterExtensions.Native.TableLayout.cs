using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double NativeAutoFitGridMinimumScale = 0.8D;

        private static (double Left, double Right) GetNativeTableOuterBorderInsets(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.BorderFrame?.Border is { } perimeter) {
                return ((GetNativeFrameSide(perimeter, top: null)?.PaintThickness ?? 0D) / 2D,
                    (GetNativeFrameSide(perimeter, top: null, right: true)?.PaintThickness ?? 0D) / 2D);
            }
            double left = 0D, right = 0D;
            int columns = GetNativeTableColumnCount(layout);
            for (int row = 0; row < layout.Rows.Count; row++) {
                int column = GetNativeTableRowStartColumn(layout, row);
                foreach (WordTableCell cell in layout.Rows[row]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) continue;
                    int span = GetNativeCellColumnSpan(cell);
                    if (column == 0) left = Math.Max(left, GetNativeTableCellBorderInset(style, row, column, right: false));
                    if (column + span == columns) right = Math.Max(right, GetNativeTableCellBorderInset(style, row, column, right: true));
                    column += span;
                }
            }
            return (left, right);
        }

        private static void ApplyNativeAutomaticTableFrameWidth(TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            if (!style.AutoFitUnspecifiedWidthToContent || style.PreferredWidth.HasValue ||
                style.AutoFitColumnPreferredWidthPoints == null || !contentWidth.HasValue) return;
            var (left, right) = GetNativeTableOuterBorderInsets(layout, style);
            double width = contentWidth.Value - style.LeftIndent - left - right -
                (style.BorderFrame == null ? 0D : 2D * style.CellSpacing);
            if (!style.AutoFitColumns) {
                // A fixed automatic frame can overflow its container. Its
                // source preferences and physical cell minima remain usable.
                double minimumFrame = style.AutoFitColumnPreferredWidthPoints.Select((preferred, column) =>
                    Math.Max(preferred ?? 0D, style.ColumnMinWidthPoints != null && column < style.ColumnMinWidthPoints.Count
                        ? style.ColumnMinWidthPoints[column] ?? 0D : 0D)).Sum();
                width = Math.Max(width, minimumFrame + Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing);
            }
            if (width > .001D) style.MaxWidth = Math.Min(style.MaxWidth ?? width, width);
        }

        private static void ApplyNativeFixedPercentageColumnPreferences(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || style.ImportedColumnWidthFractions == null) return;
            var absolute = new bool[layout.ColumnWidths.Length];
            var cached = new bool[layout.ColumnWidths.Length];
            foreach (var cell in EnumerateNativeTableCells(layout)) {
                bool hasAbsoluteWidth = TablePreferredWidth.AbsoluteCellPoints(cell.Cell).HasValue;
                bool hasCachedWidth = !GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue;
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cached.Length; column++) {
                    absolute[column] |= hasAbsoluteWidth;
                    cached[column] |= hasCachedWidth;
                }
            }
            style.AutoFitColumnPreferredWidthPoints = layout.ColumnWidths.Select((width, column) => {
                if (!absolute[column] && !cached[column] && style.ImportedColumnWidthFractions[column].HasValue)
                    return (double?)null;
                // Separate-border cached grids include the outer and intercell
                // spacing shares. Absolute cell preferences describe the box.
                double spacingShare = !absolute[column] && style.BorderFrame != null
                    ? style.CellSpacing * (cached.Length == 1 ? 2D : column == 0 || column == cached.Length - 1 ? 1.5D : 1D) : 0D;
                return (double?)Math.Max(.001D, width - spacingShare);
            }).ToList();
        }

        private static void ApplyNativeColumnWidths(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            ApplyNativeAbsoluteNoWrapMinimums(layout, style);
            if (style.AutoFitColumns) {
                style.AutoFitCellPreferredWidthPoints =
                    EnumerateNativeTableCellPositions(layout)
                        .Where(cell => TablePreferredWidth.HasAbsoluteCellPreference(cell.Cell))
                        .ToDictionary(cell => (cell.Row, cell.Column), cell => Math.Max(.001D,
                            TablePreferredWidth.AbsoluteCellPoints(cell.Cell) ??
                            layout.ColumnWidths.Skip(cell.Column).Take(cell.ColumnSpan).Sum(width => (double)width)));
            }
            if (!style.AutoFitColumns && style.ImportedColumnWidthFractions == null && !style.MaxWidth.HasValue &&
                layout.ColumnWidths.Length > 0 && layout.ColumnWidths.All(width => width >= 0F)) {
                double gapWidth = Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing;
                style.MaxWidth = Math.Max(layout.ColumnWidths.Sum(width => (double)width),
                    style.ColumnMinWidthPoints?.Sum(width => width ?? 0D) ?? 0D) + gapWidth;
                style.PreserveWidth = true;
                style.AllowWidthBeyondFrame = true;
            }
            if (style.ImportedColumnWidthFractions != null) {
                style.ColumnWidthPoints = null;
                style.ColumnWidthWeights = style.ImportedColumnWidthFractions.All(value => value.HasValue)
                    ? style.ImportedColumnWidthFractions.Select(value => value!.Value).ToList() : null;
                return;
            }

            if (style.AutoFitColumns) {
                if (style.AutoFitColumnPreferredWidthPoints != null) return;
                if (layout.ColumnWidths.Length > 0 && layout.ColumnWidths.All(width => width > 0)) {
                    ApplyNativeAutoFitGridMinimums(layout, style, contentWidth);
                }
                if (layout.ColumnWidths.Length > 0 &&
                    layout.ColumnWidths.All(width => width > 0) &&
                    (!IsNativeTableAutoFitToContents(table._tableProperties) || HasNativeTableCellWidthPreferences(layout)) &&
                    layout.ColumnWidths.Skip(1).Any(width => Math.Abs(width - layout.ColumnWidths[0]) > 0.01F)) {
                    style.ColumnWidthWeights = layout.ColumnWidths.Select(width => (double)width).ToList();
                }
                return;
            }

            style.ColumnWidthPoints = CreateNativeColumnWidthPoints(layout, style);
        }

        /// <summary>No-wrap absolute cells can wrap words, but retain their preferred allocation before other cells shrink.</summary>
        private static void ApplyNativeAbsoluteNoWrapMinimums(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (!style.AutoFitColumns) return;
            foreach (var cell in EnumerateNativeTableCellPositions(layout)) {
                double? width = TablePreferredWidth.AbsoluteCellPoints(cell.Cell);
                if (cell.Cell.WrapText || width is not > 0D) continue;
                style.ColumnMinWidthPoints ??= new List<double?>();
                while (style.ColumnMinWidthPoints.Count < layout.ColumnWidths.Length)
                    style.ColumnMinWidthPoints.Add(null);
                double borderWidth = style.BorderFrame?.Border == null ? 0D : 2D * (
                    GetNativeTableCellBorderInset(style, cell.Row, cell.Column, right: false) +
                    GetNativeTableCellBorderInset(style, cell.Row, cell.Column, right: true));
                ApplyNativeSpanMinimum(layout, style.ColumnMinWidthPoints, cell.Column, cell.ColumnSpan,
                    Math.Max(0D, width.Value + borderWidth - Math.Max(0, cell.ColumnSpan - 1) * style.CellSpacing));
            }
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
                    if (!TablePreferredWidth.HasAbsoluteCellPreference(cell.Cell))
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

        private static List<double?>? CreateNativeColumnWidthFractions(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var weights = new double[columnCount];
            var hasPercentWidth = new bool[columnCount];
            bool hasAnyPercentWidth = false;
            foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout).OrderBy(cell => cell.ColumnSpan)) {
                double? percent = GetNativeTableCellPreferredWidthPercent(cell.Cell);
                if (!percent.HasValue) {
                    continue;
                }

                double columnWeight = percent.Value / cell.ColumnSpan;
                if (cell.ColumnSpan > 1) {
                    // Individual cells constrain their own columns. A merged
                    // preference supplies only the still-unassigned share.
                    double assigned = 0D;
                    int missing = 0;
                    for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < weights.Length; column++) {
                        if (hasPercentWidth[column]) assigned += weights[column];
                        else missing++;
                    }
                    if (missing == 0) continue;
                    if (percent.Value > assigned) columnWeight = (percent.Value - assigned) / missing;
                }
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < weights.Length; columnIndex++) {
                    if (cell.ColumnSpan > 1 && hasPercentWidth[columnIndex]) continue;
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

            return weights.Select((weight, column) => hasPercentWidth[column] ? (double?)weight : null).ToList();
        }

        private static List<double?>? CreateNativeColumnWidthPoints(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || layout.ColumnWidths.Length == 0 || !layout.ColumnWidths.All(width => width >= 0)) {
                return null;
            }

            // Core fits the complete grid once, locking columns at their
            // physical minimums before redistributing the remaining width.
            // Pre-scaling here can make a narrow column invalid before that fit.
            return layout.ColumnWidths.Select((width, column) => (double?)Math.Max(Math.Max(.001D, width),
                style.ColumnMinWidthPoints != null && column < style.ColumnMinWidthPoints.Count
                    ? style.ColumnMinWidthPoints[column] ?? 0D : 0D)).ToList();
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableProperties? properties = table._tableProperties;
            if (ShouldUseNativeAutoFitTableLayout(properties, tableStyleDefaults)) {
                style.AutoFitColumns = true;
            }
            style.ImportedColumnWidthFractions = CreateNativeColumnWidthFractions(layout);

            double? cellSpacing = GetNativeTableCellSpacing(properties?.TableCellSpacing) ?? tableStyleDefaults.CellSpacing;
            if (cellSpacing.HasValue) {
                style.CellSpacing = cellSpacing.Value;
            }

            double? preferredWidth = GetNativeTablePreferredWidth(properties?.TableWidth, contentWidth) ??
                GetNativeTablePreferredWidth(tableStyleDefaults.PreferredWidth, contentWidth);
            if (preferredWidth.HasValue && !style.AutoFitColumns) {
                style.MaxWidth = preferredWidth.Value;
                style.PreserveWidth = true;
                // Fixed source grids retain their width in narrow columns and
                // beyond the page text area, including after indentation.
                style.AllowWidthBeyondFrame = true;
            } else {
                preferredWidth ??= GetNativeAutoFitGridPreferredWidth(properties, layout, contentWidth, style.CellSpacing);
                if (preferredWidth.HasValue) {
                    // An automatic table's preferred width can grow to fit an
                    // unbreakable cell; it is not a clipping boundary.
                    style.PreferredWidth = preferredWidth.Value;
                    W.TableWidth? authoredWidth = properties?.TableWidth;
                    if (!GetNativeTablePreferredWidth(authoredWidth, contentWidth).HasValue)
                        authoredWidth = tableStyleDefaults.PreferredWidth;
                    if (style.AutoFitColumns && ResolveNativeTablePreferredWidthType(authoredWidth) == W.TableWidthUnitValues.Dxa) {
                        style.AllowWidthBeyondFrame = true;
                        style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
                    }
                    // Positive spacing receives its independent perimeter and
                    // final cell-grid minimums after cell formatting is applied.
                    style.AutoFitWidthUsesContentMinimum = style.AutoFitColumns && style.CellSpacing <= 0D;
                    style.PreserveWidth = true;
                }
            }

            if (style.AutoFitColumns && IsNativeTableAutoFitToContents(properties)) {
                // A valid automatic grid can contain only zero/automatic widths.
                // Its physical content minimum still determines the table frame.
                style.AutoFitWidthUsesContentMinimum = true;
                style.AutoFitUnspecifiedWidthToContent = true;
                style.PreserveWidth = true;
                // An automatic frame derives its width from cell preferences
                // and content, independently of the stored grid extent.
                if (preferredWidth.HasValue) style.PreferredWidth = null;
                style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
            }
            if (!style.AutoFitColumns && style.ImportedColumnWidthFractions != null) {
                style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
                foreach (var cell in EnumerateNativeTableCells(layout))
                    if (!GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue)
                        for (int column = cell.Column; column < cell.Column + cell.ColumnSpan &&
                            column < style.AutoFitColumnPreferredWidthPoints.Count; column++)
                            style.AutoFitColumnPreferredWidthPoints[column] ??= Math.Max(.001D, layout.ColumnWidths[column]);
                for (int column = 0; column < style.ImportedColumnWidthFractions.Count; column++)
                    if (!style.ImportedColumnWidthFractions[column].HasValue)
                        style.AutoFitColumnPreferredWidthPoints[column] ??= Math.Max(.001D, layout.ColumnWidths[column]);
                if (ResolveNativeTablePreferredWidthType(properties?.TableWidth) == W.TableWidthUnitValues.Auto) {
                    style.PreferredWidth = null;
                    style.AutoFitUnspecifiedWidthToContent = true;
                    style.PreserveWidth = true;
                    style.AllowWidthBeyondFrame = true;
                }
            }

            // tblInd applies only to left-aligned tables. A floating coordinate
            // remains a placement coordinate even in the column fallback path.
            double? leftIndent = GetNativeTableHorizontalPositionIndent(properties?.TablePositionProperties) ??
                (MapNativeTableAlignment(ResolveNativeTableAlignment(table, tableStyleDefaults)) == PdfCore.PdfAlign.Left
                    ? GetNativeTableLeftIndent(properties?.TableIndentation) ?? tableStyleDefaults.LeftIndent
                    : null);
            if (leftIndent.HasValue) {
                style.LeftIndent = leftIndent.Value;
            }

        }

        private static List<double?> CreateNativeAutoFitColumnPreferences(TableLayout layout) {
            var result = Enumerable.Repeat<double?>(null, layout.ColumnWidths.Length).ToList();
            foreach (var cell in EnumerateNativeTableCells(layout)) {
                if (!TablePreferredWidth.HasAbsoluteCellPreference(cell.Cell)) continue;
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < result.Count; column++)
                    result[column] = Math.Max(.001D, layout.ColumnWidths[column]);
            }
            return result;
        }

        private static double? GetNativeAutoFitGridPreferredWidth(W.TableProperties? properties, TableLayout layout, double? contentWidth, double cellSpacing) {
            if (!IsNativeTableAutoFitToContents(properties) ||
                layout.ColumnWidths.Length == 0 ||
                !layout.ColumnWidths.All(width => width > 0F) ||
                !HasNativeTableCellWidthPreferences(layout)) {
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

        private static bool HasNativeTableCellWidthPreferences(TableLayout layout) =>
            EnumerateNativeTableCells(layout).Any(cell =>
                TablePreferredWidth.HasAbsoluteCellPreference(cell.Cell) ||
                GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue);

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
            ResolveNativeTablePreferredWidthType(properties?.TableWidth) == W.TableWidthUnitValues.Auto;

        private static bool IsNativeTableAutoFitLayout(W.TableProperties? properties) =>
            properties?.TableLayout?.Type?.Value != W.TableLayoutValues.Fixed;

        /// <summary>Resolves omitted defaults and Word's zero, nil and percentage preferred-width rules.</summary>
        private static W.TableWidthUnitValues ResolveNativeTablePreferredWidthType(W.TableWidth? width) =>
            TablePreferredWidth.ResolveType(width);

        private static double? GetNativeTablePreferredWidth(W.TableWidth? width, double? contentWidth) {
            W.TableWidthUnitValues type = ResolveNativeTablePreferredWidthType(width);
            if (type == W.TableWidthUnitValues.Pct && width != null) {
                double? percent = GetNativeTablePreferredWidthPercent(width);
                if (!percent.HasValue || !contentWidth.HasValue || contentWidth.Value <= 0D) {
                    return null;
                }

                return contentWidth.Value * percent.Value;
            }

            if (type != W.TableWidthUnitValues.Dxa || width == null) {
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
            if (TablePreferredWidth.ResolveCellType(width) != W.TableWidthUnitValues.Pct) {
                return null;
            }

            return GetNativeTableWidthPercent(width?.Width?.Value);
        }

        private static double? GetNativeTableWidthPercent(string? rawWidth) => TablePreferredWidth.Percentage(rawWidth);

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

        /// <summary>An absolute spacing declaration selects separate border geometry even when its value is zero.</summary>
        private static bool UsesNativeSeparateTableCellBorders(WordTable table, PdfCore.PdfTableStyle style,
            NativeTableStyleDefaults defaults) {
            if (style.CellSpacing > 0D) return true;
            W.TableCellSpacing? spacing = table._tableProperties?.TableCellSpacing;
            return GetNativeTableCellSpacing(spacing).HasValue
                ? spacing!.Type?.Value == W.TableWidthUnitValues.Dxa
                : defaults.SeparateCellBorders == true;
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
            WordTableAlignmentResolver.Read(table._tableProperties?.TableJustification, tableStyleDefaults.RightToLeft) ??
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
