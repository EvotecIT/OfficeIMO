using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static ConditionalFillData? ReadConditionalFillData(ExcelSheet? workbookSheet, object?[,] values, string?[,]? cellReferences, bool enabled) {
            if (!enabled || workbookSheet == null || cellReferences == null) {
                return null;
            }

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = workbookSheet.GetConditionalFormattingRules();
            if (rules.Count == 0) {
                return null;
            }

            var fills = new Dictionary<(int Row, int Column), string>();
            var dataBars = new Dictionary<(int Row, int Column), ConditionalDataBarCell>();
            var icons = new Dictionary<(int Row, int Column), ConditionalIconCell>();
            var hiddenValues = new HashSet<(int Row, int Column)>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "ColorScale", StringComparison.OrdinalIgnoreCase) && rule.ColorScaleColors.Count >= 2)
                .OrderByDescending(rule => rule.Priority)) {
                if (!TryGetRgb(rule.ColorScaleColors[0], out _, out _, out _) ||
                    !TryGetRgb(rule.ColorScaleColors[rule.ColorScaleColors.Count - 1], out _, out _, out _)) {
                    continue;
                }

                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                IReadOnlyList<double> candidateValues = candidates.Select(candidate => candidate.Value).ToArray();
                if (!ExcelConditionalFormatThresholds.TryCreateColorScaleEvaluator(
                        candidateValues,
                        rule.ColorScaleColors,
                        rule.ColorScaleThresholds,
                        out ExcelConditionalFormatThresholds.ColorScaleEvaluator? colorScale) ||
                    colorScale == null) {
                    continue;
                }
                foreach (var candidate in candidates) {
                    fills[(candidate.Row, candidate.Column)] = colorScale.GetRgbHex(candidate.Value);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "DataBar", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.DataBarColor))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                IReadOnlyList<double> candidateValues = candidates.Select(candidate => candidate.Value).ToArray();
                (double min, double max) = ExcelConditionalFormatThresholds.ResolveDataBarRange(candidateValues, rule.DataBarThresholds);
                // Excel paints data bars as a gradient that fades to white, so the value on
                // top stays readable. A flat fill takes the gradient's average tint.
                string barColor = rule.DataBarGradient != false ? BlendTowardWhite(rule.DataBarColor!, 0.5D) : rule.DataBarColor!;
                foreach (var candidate in candidates) {
                    (double startRatio, double ratio) = ExcelConditionalFormatThresholds.GetDataBarGeometry(candidate.Value, min, max);
                    var key = (candidate.Row, candidate.Column);
                    dataBars[key] = new ConditionalDataBarCell(barColor, startRatio, ratio);
                    if (!rule.DataBarShowValue) hiddenValues.Add(key);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.IconSet))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                IReadOnlyList<double> ruleValues = ExcelConditionalVisualEvaluator.GetIconRuleNumericValues(workbookSheet, rule.Range);
                if (ruleValues.Count == 0) ruleValues = candidates.Select(candidate => candidate.Value).ToArray();
                double min = ruleValues.Min();
                double max = ruleValues.Max();
                foreach (var candidate in candidates) {
                    if (!ExcelConditionalVisualEvaluator.TryResolveConditionalIcon(rule, candidate.Value, ruleValues, min, max, out OfficeConditionalIconKind? kind)) continue;
                    var key = (candidate.Row, candidate.Column);
                    // A higher-priority NoIcons substitution also replaces a lower-priority icon.
                    icons.Remove(key);
                    icons[key] = new ConditionalIconCell(kind ?? OfficeConditionalIconKind.None, rule.IconSetShowValue);
                    if (rule.IconSetShowValue) hiddenValues.Remove(key);
                    else hiddenValues.Add(key);
                }
            }

            AddRuleConditionalFills(workbookSheet, cellReferences, rules, fills);

            // Empty strings suppress the value even when the caller supplies custom empty-cell text.
            foreach (var key in hiddenValues) values[key.Row, key.Column] = string.Empty;
            return fills.Count == 0 && dataBars.Count == 0 && icons.Count == 0 ? null : new ConditionalFillData(fills, dataBars, icons);
        }

        /// <summary>
        /// Applies the fills of rule-based formats (cell value, text, dates, duplicates, top/bottom,
        /// average and formula rules) through the same evaluator the image exporter uses. Colour
        /// scales already filled above keep their colour.
        /// </summary>
        private static void AddRuleConditionalFills(ExcelSheet sheet, string?[,] cellReferences,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules, Dictionary<(int Row, int Column), string> fills) {
            if (!rules.Any(rule => !IsScaleBarOrIconRule(rule.Type))) return;
            int rows = cellReferences.GetLength(0), columns = cellReferences.GetLength(1);
            int firstRow = int.MaxValue, firstColumn = int.MaxValue, lastRow = 0, lastColumn = 0;
            var positions = new Dictionary<(int Row, int Column), (int Row, int Column)>();
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    string? reference = cellReferences[row, column];
                    if (string.IsNullOrWhiteSpace(reference)) continue;
                    (int sourceRow, int sourceColumn) = ExcelA1.ParseCellRef(NormalizeCellReference(reference!));
                    if (sourceRow <= 0 || sourceColumn <= 0) continue;
                    positions[(sourceRow, sourceColumn)] = (row, column);
                    firstRow = Math.Min(firstRow, sourceRow); lastRow = Math.Max(lastRow, sourceRow);
                    firstColumn = Math.Min(firstColumn, sourceColumn); lastColumn = Math.Max(lastColumn, sourceColumn);
                }
            }

            if (positions.Count == 0) return;
            var options = new ExcelImageExportOptions {
                IncludeHidden = true,
                IncludeImages = false,
                IncludeCharts = false,
                IncludeDrawingObjects = false,
                IncludeConditionalFormatting = false,
                ShowHyperlinkHints = false
            };
            if ((long)(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) > options.MaximumRenderedCells) return;
            string range = ExcelA1.CellReference(firstRow, firstColumn) + ":" + ExcelA1.CellReference(lastRow, lastColumn);
            ExcelConditionalVisualState state;
            try {
                ExcelRangeVisualInfo snapshot = ExcelRangeVisualSnapshotBuilder.Build(sheet, range, options);
                state = ExcelConditionalVisualEvaluator.Evaluate(sheet, snapshot.Cells, range, DateTime.Today, new List<OfficeImageExportDiagnostic>());
            } catch (Exception exception) when (exception is InvalidOperationException or ArgumentException or NotSupportedException) {
                return;
            }

            foreach (KeyValuePair<string, ExcelConditionalCellFormat> format in state.CellFormats) {
                string? fill = format.Value.FillColorArgb?.Trim().TrimStart('#');
                if (fill == null || (fill.Length != 6 && fill.Length != 8)) continue;
                if (fill.Length == 8) fill = fill.Substring(2);
                int separator = format.Key.IndexOf(':');
                if (separator <= 0 ||
                    !int.TryParse(format.Key.Substring(0, separator), NumberStyles.Integer, CultureInfo.InvariantCulture, out int sourceRow) ||
                    !int.TryParse(format.Key.Substring(separator + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int sourceColumn) ||
                    !positions.TryGetValue((sourceRow, sourceColumn), out var position)) continue;
                if (!fills.ContainsKey(position)) fills[position] = fill.ToUpperInvariant();
            }
        }

        private static bool IsScaleBarOrIconRule(string? type) =>
            string.Equals(type, "ColorScale", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "DataBar", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "IconSet", StringComparison.OrdinalIgnoreCase);

        private static bool IsCellReferenceInReferenceList(string cellReference, string referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            (int Row, int Col) cell = ExcelA1.ParseCellRef(NormalizeCellReference(cellReference));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (ExcelA1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    if (cell.Row >= firstRow && cell.Row <= lastRow && cell.Col >= firstColumn && cell.Col <= lastColumn) {
                        return true;
                    }
                } else {
                    (int Row, int Col) singleCell = ExcelA1.ParseCellRef(token);
                    if (singleCell.Row == cell.Row && singleCell.Col == cell.Col) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetConditionalNumericValue(object? value, out double numericValue) {
            if (value is DateTime dateTime) {
                numericValue = dateTime.ToOADate();
                return true;
            }

            if (value is IConvertible convertible) {
                try {
                    numericValue = convertible.ToDouble(CultureInfo.InvariantCulture);
                    return !double.IsNaN(numericValue) && !double.IsInfinity(numericValue);
                } catch (FormatException) {
                } catch (InvalidCastException) {
                } catch (OverflowException) {
                }
            }

            numericValue = 0D;
            return false;
        }

        private static string BlendTowardWhite(string color, double amount) {
            if (!TryGetRgb(color, out byte r, out byte g, out byte b)) return color;
            byte Blend(byte channel) => (byte)Math.Round(channel + (255 - channel) * amount);
            return "#" + Blend(r).ToString("X2", CultureInfo.InvariantCulture) + Blend(g).ToString("X2", CultureInfo.InvariantCulture) + Blend(b).ToString("X2", CultureInfo.InvariantCulture);
        }

        private static bool TryGetRgb(string value, out byte r, out byte g, out byte b) =>
            ExcelConditionalFormatThresholds.TryGetRgb(value, out r, out g, out b);

    }
}
