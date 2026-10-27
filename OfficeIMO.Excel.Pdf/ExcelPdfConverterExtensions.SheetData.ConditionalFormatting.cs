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
                foreach (var candidate in candidates) {
                    (double startRatio, double ratio) = ExcelConditionalFormatThresholds.GetDataBarGeometry(candidate.Value, min, max);
                    dataBars[(candidate.Row, candidate.Column)] = new ConditionalDataBarCell(rule.DataBarColor!, startRatio, ratio);
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

            // Empty strings suppress the value even when the caller supplies custom empty-cell text.
            foreach (var key in hiddenValues) values[key.Row, key.Column] = string.Empty;
            return fills.Count == 0 && dataBars.Count == 0 && icons.Count == 0 ? null : new ConditionalFillData(fills, dataBars, icons);
        }

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

        private static bool TryGetRgb(string value, out byte r, out byte g, out byte b) =>
            ExcelConditionalFormatThresholds.TryGetRgb(value, out r, out g, out b);

    }
}
