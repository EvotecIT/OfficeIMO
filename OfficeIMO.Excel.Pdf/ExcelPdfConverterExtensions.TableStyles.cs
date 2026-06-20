using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static PdfCore.PdfTableStyle CreateTableStyle(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup, IReadOnlyList<int> rowIndexes, int headerRowCount, ExcelCellStyleSnapshot?[,]? styles, ConditionalFillData? conditionalFills, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int columnOffset = 0, int exportedColumns = 0) {
            int exportedRows = rowIndexes.Count;
            int headerRows = Math.Min(headerRowCount, exportedRows);
            PdfCore.PdfTableStyle tableStyle = CreateBaseTableStyle(options);
            tableStyle.HeaderRowCount = headerRows;
            tableStyle.RepeatHeaderRowCount = headerRows == 0 ? null : headerRows;

            if (columnWidths != null) {
                List<double> widthWeights = columnWidths.WidthWeights.Skip(columnOffset).Take(exportedColumns).ToList();
                tableStyle.ColumnWidthWeights = widthWeights.Count == 0 ? columnWidths.WidthWeights : widthWeights;
                if (IsFitToWidth(pageSetup)) {
                    tableStyle.MaxWidth = CalculateFitToWidthMaxWidth(options, pageSetup);
                } else if (pageSetup?.Scale is uint scale && scale > 0U && scale < 100U) {
                    double approximateWidth = CalculateChunkApproximateWidth(columnWidths, columnOffset, exportedColumns);
                    tableStyle.MaxWidth = Math.Max(24D, approximateWidth * scale / 100D);
                }
            }

            if (rowHeights != null) {
                tableStyle.RowMinHeights = rowIndexes
                    .Select(row => row >= 0 && row < rowHeights.MinHeights.Count ? rowHeights.MinHeights[row] : null)
                    .ToList();
            }

            ApplyFitToHeight(tableStyle, options, pageSetup, rowIndexes, rowHeights);

            Dictionary<(int Row, int Column), PdfCore.PdfColor>? cellFills = CreateCellFills(styles, conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellFills != null) {
                tableStyle.CellFills = cellFills;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? cellDataBars = CreateCellDataBars(conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellDataBars != null) {
                tableStyle.CellDataBars = cellDataBars;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? cellIcons = CreateCellIcons(conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellIcons != null) {
                tableStyle.CellIcons = cellIcons;
                tableStyle.CellPaddings = CreateIconCellPaddings(cellIcons, tableStyle.CellPaddings);
            }

            Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? cellAlignments = CreateCellAlignments(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellAlignments != null) {
                tableStyle.CellAlignments = cellAlignments;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? cellVerticalAlignments = CreateCellVerticalAlignments(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellVerticalAlignments != null) {
                tableStyle.CellVerticalAlignments = cellVerticalAlignments;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = CreateCellBorders(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellBorders != null) {
                tableStyle.CellBorders = cellBorders;
            }

            return tableStyle;
        }

        private static PdfCore.PdfTableStyle CreateBaseTableStyle(ExcelPdfSaveOptions options) {
            PdfCore.PdfTableStyle? configuredStyle = options.PdfOptions?.HasExplicitDefaultTableStyle == true
                ? options.PdfOptions.DefaultTableStyle
                : null;
            if (configuredStyle != null) {
                return configuredStyle.Clone();
            }

            return new PdfCore.PdfTableStyle {
                CellPaddingX = 4,
                CellPaddingY = 3,
                HeaderFill = PdfCore.PdfColor.FromRgb(230, 238, 247),
                HeaderTextColor = PdfCore.PdfColor.FromRgb(31, 78, 121),
                RowStripeFill = PdfCore.PdfColor.FromRgb(248, 250, 252)
            };
        }

        private static PdfCore.PdfTableStyle CreateEmptyWorkbookTableStyle(ExcelPdfSaveOptions options) {
            PdfCore.PdfTableStyle tableStyle = CreateBaseTableStyle(options);
            tableStyle.HeaderRowCount = 0;
            tableStyle.RepeatHeaderRowCount = null;
            return tableStyle;
        }

        private static double CalculateChunkApproximateWidth(ColumnLayoutData columnWidths, int columnOffset, int exportedColumns) {
            if (exportedColumns <= 0 || columnOffset <= 0 && exportedColumns >= columnWidths.WidthWeights.Count) {
                return columnWidths.ApproximateWidthPoints;
            }

            double totalWeight = columnWidths.WidthWeights.Sum();
            double chunkWeight = columnWidths.WidthWeights.Skip(columnOffset).Take(exportedColumns).Sum();
            if (totalWeight <= 0D || chunkWeight <= 0D) {
                return columnWidths.ApproximateWidthPoints;
            }

            return columnWidths.ApproximateWidthPoints * chunkWeight / totalWeight;
        }

        private static bool IsFitToWidth(ExcelSheetPageSetup? pageSetup) {
            return pageSetup?.FitToWidth is uint fitToWidth && fitToWidth > 0U;
        }

        private static void ApplyFitToHeight(PdfCore.PdfTableStyle tableStyle, ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup, IReadOnlyList<int> rowIndexes, RowLayoutData? rowHeights) {
            if (!IsFitToHeight(pageSetup) || rowIndexes.Count == 0) {
                return;
            }

            double targetHeight = CalculateFitToHeightMaxHeight(options, pageSetup);
            List<double> currentHeights = CreateApproximateRowHeights(tableStyle, rowIndexes, rowHeights);
            double currentHeight = currentHeights.Sum();
            if (currentHeight <= targetHeight || currentHeight <= 0D) {
                return;
            }

            double scale = Math.Max(0.05D, targetHeight / currentHeight);
            tableStyle.FixedRowHeights = currentHeights
                .Select(height => (double?)Math.Max(1D, height * scale))
                .ToList();

            tableStyle.CellPaddingY *= scale;
            tableStyle.CellPaddingTop = ScaleOptional(tableStyle.CellPaddingTop, scale);
            tableStyle.CellPaddingBottom = ScaleOptional(tableStyle.CellPaddingBottom, scale);
            tableStyle.FontSize = ScaleFontSize(tableStyle.FontSize, scale);
            tableStyle.HeaderFontSize = ScaleFontSize(tableStyle.HeaderFontSize, scale);
            tableStyle.FooterFontSize = ScaleFontSize(tableStyle.FooterFontSize, scale);
            tableStyle.MinRowHeight *= scale;
            if (tableStyle.RowMinHeights != null) {
                tableStyle.RowMinHeights = tableStyle.RowMinHeights
                    .Select(height => height.HasValue ? (double?)Math.Max(1D, height.Value * scale) : null)
                    .ToList();
            }
        }

        private static bool IsFitToHeight(ExcelSheetPageSetup? pageSetup) {
            return pageSetup?.FitToHeight is uint fitToHeight && fitToHeight > 0U;
        }

        private static double CalculateFitToHeightMaxHeight(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            PdfCore.PageSize pageSize = GetEffectivePageSize(options, pageSetup);
            PdfCore.PageMargins margins = GetEffectiveMargins(options, pageSetup);
            double pageContentHeight = Math.Max(24D, pageSize.Height - margins.Top - margins.Bottom);
            uint fitToHeight = pageSetup?.FitToHeight ?? 1U;
            return pageContentHeight * Math.Max(1U, fitToHeight);
        }

        private static List<double> CreateApproximateRowHeights(PdfCore.PdfTableStyle tableStyle, IReadOnlyList<int> rowIndexes, RowLayoutData? rowHeights) {
            var heights = new List<double>(rowIndexes.Count);
            double fallbackHeight = GetDefaultApproximateRowHeight(tableStyle);
            for (int i = 0; i < rowIndexes.Count; i++) {
                int row = rowIndexes[i];
                double? configuredHeight = rowHeights != null && row >= 0 && row < rowHeights.MinHeights.Count
                    ? rowHeights.MinHeights[row]
                    : null;
                heights.Add(Math.Max(1D, configuredHeight ?? fallbackHeight));
            }

            return heights;
        }

        private static double GetDefaultApproximateRowHeight(PdfCore.PdfTableStyle tableStyle) {
            double fontSize = tableStyle.FontSize ?? 10D;
            double lineHeight = tableStyle.LineHeight.HasValue ? fontSize * tableStyle.LineHeight.Value : fontSize * 1.4D;
            return Math.Max(tableStyle.MinRowHeight, lineHeight + tableStyle.CellPaddingY * 2D);
        }

        private static double? ScaleOptional(double? value, double scale) {
            return value.HasValue ? Math.Max(0D, value.Value * scale) : null;
        }

        private static double? ScaleFontSize(double? value, double scale) {
            return value.HasValue ? Math.Max(0.1D, value.Value * scale) : value;
        }

        private static double CalculateFitToWidthMaxWidth(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            PdfCore.PageSize pageSize = GetEffectivePageSize(options, pageSetup);
            PdfCore.PageMargins margins = GetEffectiveMargins(options, pageSetup);
            return Math.Max(24D, pageSize.Width - margins.Left - margins.Right);
        }

        private static PdfCore.PageSize GetEffectivePageSize(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            PdfCore.PageSize pageSize = options.PageSize ?? PdfCore.PageSizes.Letter;
            if (pageSetup?.Orientation == ExcelPageOrientation.Landscape) {
                return pageSize.Landscape();
            }

            if (pageSetup?.Orientation == ExcelPageOrientation.Portrait) {
                return pageSize.Portrait();
            }

            return pageSize;
        }

        private static PdfCore.PageMargins GetEffectiveMargins(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            if (options.Margins.HasValue) {
                return options.Margins.Value;
            }

            if (pageSetup?.Margins != null) {
                return ToPdfMargins(pageSetup.Margins);
            }

            return PdfCore.PageMargins.Normal;
        }

        private static ExcelCellStyleSnapshot? GetCellStyle(ExcelCellStyleSnapshot?[,]? styles, int row, int column) {
            if (styles == null || row >= styles.GetLength(0) || column >= styles.GetLength(1)) {
                return null;
            }

            return styles[row, column];
        }

        private static ExcelHyperlinkSnapshot? GetHyperlink(ExcelHyperlinkSnapshot?[,]? hyperlinks, int row, int column) {
            if (hyperlinks == null || row >= hyperlinks.GetLength(0) || column >= hyperlinks.GetLength(1)) {
                return null;
            }

            return hyperlinks[row, column];
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfColor>? CreateCellFills(ExcelCellStyleSnapshot?[,]? styles, ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null && conditionalFills == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfColor>? fills = null;
            if (styles != null) {
                int columns = Math.Min(columnEnd, styles.GetLength(1));
                for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                    int row = rowIndexes[localRow];
                    if (row < 0 || row >= styles.GetLength(0)) {
                        continue;
                    }

                    for (int column = columnOffset; column < columns; column++) {
                        PdfCore.PdfColor? fill = ToPdfColor(styles[row, column]?.FillColorHex);
                        if (fill.HasValue) {
                            fills ??= new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
                            fills[(localRow, column - columnOffset)] = fill.Value;
                        }
                    }
                }
            }

            if (conditionalFills != null) {
                foreach (KeyValuePair<(int Row, int Column), string> conditionalFill in conditionalFills.FillColors) {
                    int localRow = FindLocalRowIndex(rowIndexes, conditionalFill.Key.Row);
                    if (localRow < 0 ||
                        conditionalFill.Key.Column < columnOffset ||
                        conditionalFill.Key.Column >= columnEnd) {
                        continue;
                    }

                    PdfCore.PdfColor? fill = ToPdfColor(conditionalFill.Value);
                    if (fill.HasValue) {
                        fills ??= new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
                        fills[(localRow, conditionalFill.Key.Column - columnOffset)] = fill.Value;
                    }
                }
            }

            return fills;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? CreateCellDataBars(ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (conditionalFills == null || conditionalFills.DataBars.Count == 0) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? dataBars = null;
            foreach (KeyValuePair<(int Row, int Column), ConditionalDataBarCell> conditionalDataBar in conditionalFills.DataBars) {
                int localRow = FindLocalRowIndex(rowIndexes, conditionalDataBar.Key.Row);
                if (localRow < 0 ||
                    conditionalDataBar.Key.Column < columnOffset ||
                    conditionalDataBar.Key.Column >= columnEnd) {
                    continue;
                }

                PdfCore.PdfColor? fill = ToPdfColor(conditionalDataBar.Value.Color);
                if (fill.HasValue) {
                    dataBars ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>();
                    dataBars[(localRow, conditionalDataBar.Key.Column - columnOffset)] = new PdfCore.PdfCellDataBar {
                        Color = fill.Value,
                        StartRatio = conditionalDataBar.Value.StartRatio,
                        Ratio = conditionalDataBar.Value.Ratio
                    };
                }
            }

            return dataBars;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? CreateCellIcons(ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (conditionalFills == null || conditionalFills.Icons.Count == 0) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? icons = null;
            foreach (KeyValuePair<(int Row, int Column), ConditionalIconCell> conditionalIcon in conditionalFills.Icons) {
                int localRow = FindLocalRowIndex(rowIndexes, conditionalIcon.Key.Row);
                if (localRow < 0 ||
                    conditionalIcon.Key.Column < columnOffset ||
                    conditionalIcon.Key.Column >= columnEnd) {
                    continue;
                }

                icons ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>();
                icons[(localRow, conditionalIcon.Key.Column - columnOffset)] = new PdfCore.PdfCellIcon {
                    Kind = conditionalIcon.Value.Kind,
                    Color = conditionalIcon.Value.Color,
                    Size = 8D
                };
            }

            return icons;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> CreateIconCellPaddings(IReadOnlyDictionary<(int Row, int Column), PdfCore.PdfCellIcon> icons, Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>? existingPaddings) {
            var paddings = existingPaddings == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(existingPaddings);

            foreach (KeyValuePair<(int Row, int Column), PdfCore.PdfCellIcon> icon in icons) {
                if (!paddings.TryGetValue(icon.Key, out PdfCore.PdfCellPadding? padding)) {
                    padding = new PdfCore.PdfCellPadding();
                } else {
                    padding = padding.Clone();
                }

                double requiredLeftPadding = icon.Value.Size + 8D;
                padding.Left = Math.Max(padding.Left ?? 0D, requiredLeftPadding);
                paddings[icon.Key] = padding;
            }

            return paddings;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? CreateCellAlignments(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? alignments = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfColumnAlign? alignment = ToPdfHorizontalAlignment(styles[row, column]?.HorizontalAlignment);
                    if (alignment.HasValue) {
                        alignments ??= new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
                        alignments[(localRow, column - columnOffset)] = alignment.Value;
                    }
                }
            }

            return alignments;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? CreateCellVerticalAlignments(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? alignments = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfCellVerticalAlign? alignment = ToPdfVerticalAlignment(styles[row, column]?.VerticalAlignment);
                    if (alignment.HasValue) {
                        alignments ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
                        alignments[(localRow, column - columnOffset)] = alignment.Value;
                    }
                }
            }

            return alignments;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateCellBorders(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? borders = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfCellBorder? border = ToPdfCellBorder(styles[row, column]?.Border);
                    if (border != null) {
                        borders ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
                        borders[(localRow, column - columnOffset)] = border;
                    }
                }
            }

            return borders;
        }

        private static int FindLocalRowIndex(IReadOnlyList<int> rowIndexes, int row) {
            for (int index = 0; index < rowIndexes.Count; index++) {
                if (rowIndexes[index] == row) {
                    return index;
                }
            }

            return -1;
        }

        private static PdfCore.PdfColumnAlign? ToPdfHorizontalAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToLowerInvariant()) {
                case "left":
                    return PdfCore.PdfColumnAlign.Left;
                case "center":
                case "centercontinuous":
                    return PdfCore.PdfColumnAlign.Center;
                case "right":
                    return PdfCore.PdfColumnAlign.Right;
                default:
                    return null;
            }
        }

        private static PdfCore.PdfCellVerticalAlign? ToPdfVerticalAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToLowerInvariant()) {
                case "top":
                    return PdfCore.PdfCellVerticalAlign.Top;
                case "center":
                    return PdfCore.PdfCellVerticalAlign.Middle;
                case "bottom":
                    return PdfCore.PdfCellVerticalAlign.Bottom;
                default:
                    return null;
            }
        }

        private static PdfCore.PdfCellBorder? ToPdfCellBorder(ExcelCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            PdfCore.PdfCellBorderSide? left = ToPdfCellBorderSide(border.Left);
            PdfCore.PdfCellBorderSide? right = ToPdfCellBorderSide(border.Right);
            PdfCore.PdfCellBorderSide? top = ToPdfCellBorderSide(border.Top);
            PdfCore.PdfCellBorderSide? bottom = ToPdfCellBorderSide(border.Bottom);
            PdfCore.PdfCellBorderSide? diagonal = ToPdfCellBorderSide(border.Diagonal);
            bool hasDiagonalUp = border.DiagonalUp && diagonal != null;
            bool hasDiagonalDown = border.DiagonalDown && diagonal != null;
            if (left == null && right == null && top == null && bottom == null && !hasDiagonalUp && !hasDiagonalDown) {
                return null;
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                TopBorder = top,
                RightBorder = right,
                BottomBorder = bottom,
                LeftBorder = left,
                DiagonalUp = hasDiagonalUp,
                DiagonalDown = hasDiagonalDown,
                DiagonalUpBorder = hasDiagonalUp ? diagonal : null,
                DiagonalDownBorder = hasDiagonalDown ? diagonal : null
            };
        }

        private static PdfCore.PdfCellBorderSide? ToPdfCellBorderSide(ExcelBorderSideSnapshot? side) {
            if (side == null) {
                return null;
            }

            double width = ToPdfBorderWidth(side.Style);
            if (width <= 0) {
                return null;
            }

            return new PdfCore.PdfCellBorderSide {
                Color = ToPdfColor(side.ColorArgb) ?? PdfCore.PdfColor.FromRgb(0, 0, 0),
                Width = width,
                DashStyle = ToPdfBorderDashStyle(side.Style),
                LineStyle = ToPdfBorderLineStyle(side.Style)
            };
        }

        private static double ToPdfBorderWidth(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return 0D;
            }

            switch (style!.Trim().ToLowerInvariant()) {
                case "none":
                    return 0D;
                case "hair":
                    return 0.25D;
                case "medium":
                case "mediumdashdot":
                case "mediumdashdotdot":
                case "mediumdashed":
                    return 1.25D;
                case "thick":
                case "double":
                    return 2D;
                default:
                    return 0.5D;
            }
        }

        private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToPdfBorderDashStyle(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            }

            switch (style!.Trim().ToLowerInvariant()) {
                case "dashed":
                case "mediumdashed":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
                case "dotted":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
                case "dashdot":
                case "dashdotdot":
                case "mediumdashdot":
                case "mediumdashdotdot":
                case "slantdashdot":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.DashDot;
                default:
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            }
        }

        private static PdfCore.PdfCellBorderLineStyle ToPdfBorderLineStyle(string? style) {
            string? normalized = style?.Trim();
            return !string.IsNullOrWhiteSpace(normalized) &&
                   string.Equals(normalized, "double", StringComparison.OrdinalIgnoreCase)
                ? PdfCore.PdfCellBorderLineStyle.TwoLine
                : PdfCore.PdfCellBorderLineStyle.Standard;
        }

        private static PdfCore.PdfColor? ToPdfColor(string? hex) {
            if (string.IsNullOrWhiteSpace(hex)) {
                return null;
            }

            string value = hex!.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }

            if (value.Length == 8) {
                value = value.Substring(2);
            }

            if (value.Length != 6 ||
                !byte.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte r) ||
                !byte.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte g) ||
                !byte.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte b)) {
                return null;
            }

            return PdfCore.PdfColor.FromRgb(r, g, b);
        }

    }
}
