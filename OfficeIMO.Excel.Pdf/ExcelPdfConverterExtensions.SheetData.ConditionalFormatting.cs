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

                int iconCount = GetExcelIconSetCount(rule.IconSet!);
                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    int bucket = GetExcelIconSetBucket(candidate.Value, min, max, iconCount);
                    if (rule.IconSetReverse) {
                        bucket = iconCount - 1 - bucket;
                    }

                    icons[(candidate.Row, candidate.Column)] = MapExcelIconSetCell(rule.IconSet!, bucket, iconCount);
                }
            }

            return fills.Count == 0 && dataBars.Count == 0 && icons.Count == 0 ? null : new ConditionalFillData(fills, dataBars, icons);
        }

        private static int GetExcelIconSetCount(string iconSet) {
            if (iconSet.StartsWith("Three", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("3", StringComparison.Ordinal)) {
                return 3;
            }

            if (iconSet.StartsWith("Four", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("4", StringComparison.Ordinal)) {
                return 4;
            }

            return 5;
        }

        private static int GetExcelIconSetBucket(double value, double min, double max, int iconCount) {
            if (iconCount <= 1 || max <= min) {
                return iconCount - 1;
            }

            double ratio = Math.Max(0D, Math.Min(1D, (value - min) / (max - min)));
            return Math.Max(0, Math.Min(iconCount - 1, (int)Math.Floor(ratio * iconCount)));
        }

        private static ConditionalIconCell MapExcelIconSetCell(string iconSet, int bucket, int iconCount) {
            string normalized = iconSet.ToLowerInvariant();
            bool trafficLights = normalized.IndexOf("traffic", StringComparison.Ordinal) >= 0;
            bool arrows = normalized.IndexOf("arrow", StringComparison.Ordinal) >= 0;
            bool symbols = normalized.IndexOf("symbol", StringComparison.Ordinal) >= 0 || normalized.IndexOf("sign", StringComparison.Ordinal) >= 0 || normalized.IndexOf("indicator", StringComparison.Ordinal) >= 0;

            if (trafficLights) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
            }

            if (arrows) {
                if (bucket == 0) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleDown, PdfCore.PdfColor.FromRgb(192, 80, 77));
                }

                if (bucket >= iconCount - 1) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(99, 155, 71));
                }

                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleRight, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            if (symbols && bucket == 0) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Diamond, PdfCore.PdfColor.FromRgb(192, 80, 77));
            }

            if (symbols && bucket >= iconCount - 1) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, PdfCore.PdfColor.FromRgb(99, 155, 71));
            }

            if (symbols) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
        }

        private static PdfCore.PdfColor GetExcelIconBucketColor(int bucket, int iconCount) {
            if (bucket <= 0) {
                return PdfCore.PdfColor.FromRgb(192, 80, 77);
            }

            if (bucket >= iconCount - 1) {
                return PdfCore.PdfColor.FromRgb(99, 155, 71);
            }

            return PdfCore.PdfColor.FromRgb(255, 192, 0);
        }

        private static bool IsCellReferenceInReferenceList(string cellReference, string referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            (int Row, int Col) cell = A1.ParseCellRef(NormalizeCellReference(cellReference));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    if (cell.Row >= firstRow && cell.Row <= lastRow && cell.Col >= firstColumn && cell.Col <= lastColumn) {
                        return true;
                    }
                } else {
                    (int Row, int Col) singleCell = A1.ParseCellRef(token);
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
