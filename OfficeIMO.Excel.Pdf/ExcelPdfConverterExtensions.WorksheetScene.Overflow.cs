using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>
        /// Cell text that Excel paints outside its own cell: the text box spans the empty
        /// neighbours the text reaches. Coordinates are unscaled points from the table's left edge.
        /// </summary>
        private sealed class WorksheetTextOverflow {
            public WorksheetTextOverflow(int row, int column, double x, double width, PdfCore.PdfColumnAlign alignment) {
                Row = row;
                Column = column;
                X = x;
                Width = width;
                Alignment = alignment;
            }

            public int Row { get; }
            public int Column { get; }
            public double X { get; }
            public double Width { get; }
            public PdfCore.PdfColumnAlign Alignment { get; }
        }

        private sealed class WorksheetSceneTextPlan {
            public Dictionary<(int Row, int Column), string> TextOverrides { get; } = new();
            public List<WorksheetTextOverflow> Overflows { get; } = new();
            public double OverflowRight { get; set; }
        }

        /// <summary>
        /// Applies Excel's display rules that depend on the column width: text spills into
        /// empty neighbours, General numbers round to fit, and other numbers and dates that
        /// do not fit show "####".
        /// </summary>
        private static WorksheetSceneTextPlan CreateWorksheetSceneTextPlan(
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            SheetExportData data,
            IReadOnlyList<double> columnWidths,
            double virtualRightRoom,
            ExcelDateSystem dateSystem,
            string emptyCellText) {
            var result = new WorksheetSceneTextPlan();
            WorksheetTextMeasurer? measurer = plan.TextMeasurer;
            if (measurer == null) return result;
            object?[,] values = data.Values;
            int rows = values.GetLength(0);
            int columns = Math.Min(values.GetLength(1), columnWidths.Count);
            var columnLeft = new double[columns + 1];
            for (int column = 0; column < columns; column++) columnLeft[column + 1] = columnLeft[column] + columnWidths[column];

            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    object? value = values[row, column];
                    if (value == null || columnWidths[column] <= 0D) continue;
                    if (data.MergedCells?.GetRegion(row, column) != null) continue;
                    ExcelCellStyleInfo? style = GetCellStyle(data.Styles, row, column);
                    if (style?.WrapText == true || style?.ShrinkToFit == true || style?.TextRotation is int rotation && rotation != 0) continue;
                    if (data.ConditionalFills?.Icons.ContainsKey((row, column)) == true) continue;
                    double available = columnWidths[column] - 2D * WorksheetCellPaddingX;
                    if (value is string text) {
                        if (text.Length == 0 || text.IndexOf('\n') >= 0) continue;
                        string displayed = FormatCellValue(value, style, emptyCellText, dateSystem);
                        double textWidth = measurer.Measure(displayed, style) + GetIndentWidth(style, measurer);
                        if (textWidth <= available) continue;
                        TryAddTextOverflow(result, data, row, column, columns, columnLeft, textWidth - available,
                            virtualRightRoom, style);
                        continue;
                    }

                    if (value is bool || !(value is DateTime || TryGetDouble(value, out _))) continue;
                    // "####" is only decided with the workbook font's own metrics; with a
                    // substitute font the column units are estimates and the value is kept.
                    if (!plan.Geometry.MaximumDigitWidthPoints.HasValue) continue;
                    // Excel's column width includes five pixels of margin around the text.
                    double numberRoom = columnWidths[column] - 5D * 0.75D;
                    string formatted = FormatCellValue(value, style, emptyCellText, dateSystem);
                    if (measurer.Measure(formatted, style) <= numberRoom + 0.01D) continue;
                    result.TextOverrides[(row, column)] = FitNumericText(value, style, formatted, numberRoom, measurer);
                }
            }

            return result;
        }

        private static double GetIndentWidth(ExcelCellStyleInfo? style, WorksheetTextMeasurer measurer) =>
            style?.TextIndent is uint indent && indent > 0U ? indent * 3D * measurer.Measure(" ", style) : 0D;

        private static void TryAddTextOverflow(
            WorksheetSceneTextPlan result,
            SheetExportData data,
            int row,
            int column,
            int columns,
            double[] columnLeft,
            double needed,
            double virtualRightRoom,
            ExcelCellStyleInfo? style) {
            PdfCore.PdfColumnAlign? alignment = GetOverflowAlignment(style?.HorizontalAlignment);
            if (!alignment.HasValue) return;
            double cellLeft = columnLeft[column];
            double cellRight = columnLeft[column + 1];
            double left = cellLeft;
            double right = cellRight;
            switch (alignment.Value) {
                case PdfCore.PdfColumnAlign.Left:
                    right = cellRight + MeasureOverflowRoom(data, row, column, columns, columnLeft, +1, needed, virtualRightRoom);
                    break;
                case PdfCore.PdfColumnAlign.Right:
                    left = cellLeft - MeasureOverflowRoom(data, row, column, columns, columnLeft, -1, needed, 0D);
                    break;
                default:
                    // Excel keeps centred text centred on its own cell and spills both ways.
                    double half = needed / 2D;
                    double extension = Math.Min(
                        MeasureOverflowRoom(data, row, column, columns, columnLeft, -1, half, 0D),
                        MeasureOverflowRoom(data, row, column, columns, columnLeft, +1, half, virtualRightRoom));
                    left = cellLeft - extension;
                    right = cellRight + extension;
                    break;
            }

            if (right - left <= cellRight - cellLeft + 0.01D) return;
            result.TextOverrides[(row, column)] = string.Empty;
            result.Overflows.Add(new WorksheetTextOverflow(row, column, left, right - left, alignment.Value));
            result.OverflowRight = Math.Max(result.OverflowRight, right);
        }

        /// <summary>Room in the empty cells next to a cell, up to what the text needs.</summary>
        private static double MeasureOverflowRoom(
            SheetExportData data,
            int row,
            int column,
            int columns,
            double[] columnLeft,
            int direction,
            double needed,
            double virtualRoom) {
            double room = 0D;
            for (int neighbour = column + direction; neighbour >= 0 && neighbour < columns && room < needed; neighbour += direction) {
                if (!IsEmptyForOverflow(data, row, neighbour)) return room;
                room += columnLeft[neighbour + 1] - columnLeft[neighbour];
            }

            bool reachedRightEdge = direction > 0;
            for (int neighbour = column + 1; reachedRightEdge && neighbour < columns; neighbour++) {
                if (!IsEmptyForOverflow(data, row, neighbour)) reachedRightEdge = false;
            }

            if (room < needed && reachedRightEdge) room += Math.Min(virtualRoom, needed - room);
            return Math.Min(room, needed);
        }

        private static bool IsEmptyForOverflow(SheetExportData data, int row, int column) {
            object? value = data.Values[row, column];
            if (value is string text ? text.Length > 0 : value != null) return false;
            if (data.MergedCells?.GetRegion(row, column) != null) return false;
            if (data.ConditionalFills?.DataBars.ContainsKey((row, column)) == true ||
                data.ConditionalFills?.Icons.ContainsKey((row, column)) == true) {
                return false;
            }

            return true;
        }

        private static PdfCore.PdfColumnAlign? GetOverflowAlignment(string? alignment) {
            switch (alignment?.Trim().ToLowerInvariant()) {
                case null:
                case "":
                case "general":
                case "left":
                    return PdfCore.PdfColumnAlign.Left;
                case "right":
                    return PdfCore.PdfColumnAlign.Right;
                case "center":
                    return PdfCore.PdfColumnAlign.Center;
                default:
                    // Fill, justify, distributed and centre-across-selection never spill.
                    return null;
            }
        }

        /// <summary>
        /// Numbers never spill. Excel rounds a General number to the column width and
        /// fills any other number or date that does not fit with "#".
        /// </summary>
        private static string FitNumericText(object value, ExcelCellStyleInfo? style, string formatted, double available, WorksheetTextMeasurer measurer) {
            string? formatCode = style?.NumberFormatCode;
            bool general = string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase);
            if (general && value is not DateTime && TryGetDouble(value, out double number)) {
                for (int characters = Math.Min(11, formatted.Length - 1); characters >= 1; characters--) {
                    string? candidate = FormatGeneralNumber(number, characters);
                    if (candidate != null && measurer.Measure(candidate, style) <= available) return candidate;
                }
            }

            double hashWidth = Math.Max(0.1D, measurer.Measure("#", style));
            int count = Math.Max(1, (int)Math.Floor(available / hashWidth));
            return new string('#', count);
        }

        /// <summary>Excel's General display of a number within a character budget, or null.</summary>
        private static string? FormatGeneralNumber(double value, int maximumCharacters) {
            if (double.IsNaN(value) || double.IsInfinity(value)) return null;
            if (value == 0D) return "0";
            double magnitude = Math.Abs(value);
            int sign = value < 0D ? 1 : 0;
            if (magnitude < 1E11 && magnitude >= 1E-9) {
                int integerDigits = magnitude >= 1D ? (int)Math.Floor(Math.Log10(magnitude)) + 1 : 1;
                int decimals = maximumCharacters - sign - integerDigits - 1;
                if (decimals >= 0 || maximumCharacters - sign >= integerDigits) {
                    string text = Math.Round(value, Math.Max(0, decimals), MidpointRounding.AwayFromZero)
                        .ToString("0." + new string('#', Math.Max(0, decimals)), CultureInfo.InvariantCulture);
                    if (text.Length <= maximumCharacters && text != "0" && text != "-0") return text;
                }
            }

            for (int digits = Math.Min(5, maximumCharacters); digits >= 0; digits--) {
                string format = digits == 0 ? "0E+00" : "0." + new string('#', digits) + "E+00";
                string scientific = value.ToString(format, CultureInfo.InvariantCulture);
                if (scientific.Length <= maximumCharacters) return scientific;
            }

            return null;
        }

        /// <summary>Paints spilled cell text over the cell grid, clipped to the empty neighbours it reaches.</summary>
        private static void AddWorksheetSceneTextOverflows(
            PdfCore.PdfPageCanvas canvas,
            WorksheetPdfExportPlan plan,
            SheetExportData data,
            WorksheetSceneTextPlan textPlan,
            PdfCore.PdfTableStyle sceneStyle,
            IReadOnlyList<double> rowHeights,
            ExcelPdfExportOptions options,
            IReadOnlyDictionary<string, string> sheetDestinations,
            IReadOnlyDictionary<string, string> cellDestinations,
            PdfCore.PdfStandardFont defaultFontFamily,
            double tableX,
            double tableY,
            double scale,
            ExcelDateSystem dateSystem) {
            if (textPlan.Overflows.Count == 0) return;
            var rowTop = new double[rowHeights.Count + 1];
            for (int row = 0; row < rowHeights.Count; row++) rowTop[row + 1] = rowTop[row] + rowHeights[row];
            foreach (WorksheetTextOverflow overflow in textPlan.Overflows) {
                if (overflow.Row >= rowHeights.Count || rowHeights[overflow.Row] <= 0D) continue;
                object? value = data.Values[overflow.Row, overflow.Column];
                ExcelCellStyleInfo? cellStyle = GetCellStyle(data.Styles, overflow.Row, overflow.Column);
                StructuredTableCellVisual? tableVisual = GetStructuredTableCellVisual(data.StructuredTables, data.CellReferences, overflow.Row, overflow.Column);
                ExcelHyperlinkInfo? hyperlink = GetHyperlink(data.Hyperlinks, overflow.Row, overflow.Column);
                string text = FormatCellValue(value, cellStyle, options.EmptyCellText, dateSystem);
                PdfCore.PdfTableCell cell = CreatePdfCell(text, cellStyle, tableVisual, hyperlink, span: null, sheetDestinations, cellDestinations,
                    plan.SheetName, cellDestinationName: null, cellImages: null, defaultFontFamily, scale, preserveWorksheetNoWrap: true, formatColor: null);
                PdfCore.PdfCellVerticalAlign verticalAlignment =
                    sceneStyle.CellVerticalAlignments != null && sceneStyle.CellVerticalAlignments.TryGetValue((overflow.Row, overflow.Column), out PdfCore.PdfCellVerticalAlign authored)
                        ? authored
                        : PdfCore.PdfCellVerticalAlign.Bottom;
                double width = overflow.Width * scale;
                double height = rowHeights[overflow.Row] * scale;
                var style = new PdfCore.PdfTableStyle {
                    BorderColor = null,
                    HeaderFill = null,
                    FooterFill = null,
                    RowStripeFill = null,
                    HeaderRowCount = 0,
                    HeaderBold = false,
                    RightAlignNumeric = false,
                    TextColor = sceneStyle.TextColor,
                    FontSize = sceneStyle.FontSize,
                    HeaderFontSize = sceneStyle.FontSize,
                    FooterFontSize = sceneStyle.FontSize,
                    LineHeight = 1D,
                    CellPaddingX = WorksheetCellPaddingX * scale,
                    CellPaddingY = 0D,
                    ColumnWidthPoints = new List<double?> { width },
                    FixedRowHeights = new List<double?> { height },
                    Alignments = new List<PdfCore.PdfColumnAlign> { overflow.Alignment },
                    VerticalAlignments = new List<PdfCore.PdfCellVerticalAlign> { verticalAlignment },
                    MinRowHeight = 0D,
                    PreserveWidth = true,
                    SpacingBefore = 0D,
                    SpacingAfter = 0D,
                    AlternativeText = "Worksheet " + plan.SheetName + " cell text"
                };
                if (sceneStyle.CellPaddings != null && sceneStyle.CellPaddings.TryGetValue((overflow.Row, overflow.Column), out PdfCore.PdfCellPadding? padding)) {
                    style.CellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> { [(0, 0)] = padding };
                }

                canvas.Table(new[] { new[] { cell } }, tableX + overflow.X * scale, tableY + rowTop[overflow.Row] * scale, width, height, style);
            }
        }
    }
}
