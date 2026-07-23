using System.Globalization;

namespace OfficeIMO.Excel {
    internal static class ExcelConditionalFormatThresholds {
        internal static bool TryNormalizeArgb(string? value, out string? argb) {
            argb = null;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string hex = value!.Trim().TrimStart('#');
            if (hex.Length == 6) {
                hex = "FF" + hex;
            } else if (hex.Length != 8) {
                return false;
            }

            for (int i = 0; i < hex.Length; i++) {
                char ch = hex[i];
                bool isHex = (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex) {
                    return false;
                }
            }

            argb = hex.ToUpperInvariant();
            return true;
        }

        internal static bool TryGetRgb(string value, out byte red, out byte green, out byte blue) {
            red = green = blue = 0;
            if (!TryNormalizeArgb(value, out string? argb) || argb == null) {
                return false;
            }

            return byte.TryParse(argb.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out red) &&
                byte.TryParse(argb.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out green) &&
                byte.TryParse(argb.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out blue);
        }

        internal static string InterpolateRgbHex(byte startR, byte startG, byte startB, byte endR, byte endG, byte endB, double ratio) {
            byte red = InterpolateByte(startR, endR, ratio);
            byte green = InterpolateByte(startG, endG, ratio);
            byte blue = InterpolateByte(startB, endB, ratio);
            return red.ToString("X2", CultureInfo.InvariantCulture) +
                green.ToString("X2", CultureInfo.InvariantCulture) +
                blue.ToString("X2", CultureInfo.InvariantCulture);
        }

        internal static (double Min, double Max) ResolveDataBarRange(IReadOnlyList<double> values, IReadOnlyList<ExcelConditionalFormatThreshold> thresholds) {
            double min = values.Min();
            double max = values.Max();
            double lower = thresholds.Count > 0 ? ResolveThresholdValue(thresholds[0], values, min, max, min) : min;
            double upper = thresholds.Count > 1 ? ResolveThresholdValue(thresholds[1], values, min, max, max) : max;
            return upper < lower ? (upper, lower) : (lower, upper);
        }

        internal static bool HasUnsupportedFormulaThresholds(IReadOnlyList<ExcelConditionalFormatThreshold> thresholds) {
            foreach (ExcelConditionalFormatThreshold threshold in thresholds) {
                if (string.Equals(threshold.Type, "formula", StringComparison.OrdinalIgnoreCase) &&
                    !TryParseThresholdNumber(threshold.Value, out _)) {
                    return true;
                }
            }

            return false;
        }

        internal static bool TryCreateColorScaleEvaluator(
            IReadOnlyList<double> values,
            IReadOnlyList<string> colors,
            IReadOnlyList<ExcelConditionalFormatThreshold> thresholds,
            out ColorScaleEvaluator? evaluator) {
            evaluator = null;
            if (values.Count == 0 || (colors.Count != 2 && colors.Count != 3)) {
                return false;
            }

            double min = values.Min();
            double max = values.Max();
            int stopCount = colors.Count;
            IReadOnlyList<double>? sortedValues = thresholds.Any(threshold =>
                string.Equals(threshold.Type, "percentile", StringComparison.OrdinalIgnoreCase))
                ? values.OrderBy(item => item).ToArray()
                : null;
            var stops = new List<ColorStop>(stopCount);
            for (int i = 0; i < stopCount; i++) {
                if (!TryGetRgb(colors[i], out byte red, out byte green, out byte blue)) {
                    return false;
                }

                double fallback = stopCount == 1 || max <= min
                    ? min
                    : min + ((max - min) * i / (stopCount - 1));
                double thresholdValue = i < thresholds.Count
                    ? ResolveThresholdValue(thresholds[i], values, min, max, fallback, sortedValues)
                    : fallback;
                stops.Add(new ColorStop(thresholdValue, red, green, blue));
            }

            evaluator = new ColorScaleEvaluator(stops);
            return true;
        }

        internal static (double StartRatio, double Ratio) GetDataBarGeometry(double value, double min, double max) {
            if (max <= min) {
                return (0D, 1D);
            }

            if (min < 0D && max > 0D) {
                double range = max - min;
                double zeroRatio = Math.Max(0D, Math.Min(1D, -min / range));
                if (value >= 0D) {
                    return (zeroRatio, Math.Max(0D, Math.Min(1D - zeroRatio, value / range)));
                }

                double ratio = Math.Max(0D, Math.Min(zeroRatio, -value / range));
                return (zeroRatio - ratio, ratio);
            }

            if (max <= 0D) {
                double maxMagnitude = Math.Max(Math.Abs(min), Math.Abs(max));
                double ratio = maxMagnitude <= 0D ? 0D : Math.Max(0D, Math.Min(1D, Math.Abs(value) / maxMagnitude));
                return (1D - ratio, ratio);
            }

            double positiveRatio = Math.Max(0D, Math.Min(1D, (value - min) / (max - min)));
            return (0D, positiveRatio);
        }

        private static double ResolveThresholdValue(
            ExcelConditionalFormatThreshold threshold,
            IReadOnlyList<double> values,
            double min,
            double max,
            double fallback,
            IReadOnlyList<double>? sortedValues = null) {
            string type = threshold.Type ?? string.Empty;
            if (string.Equals(type, "min", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "minimum", StringComparison.OrdinalIgnoreCase)) {
                return min;
            }

            if (string.Equals(type, "max", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "maximum", StringComparison.OrdinalIgnoreCase)) {
                return max;
            }

            if (string.Equals(type, "num", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "number", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "formula", StringComparison.OrdinalIgnoreCase)) {
                return TryParseThresholdNumber(threshold.Value, out double number) ? number : fallback;
            }

            if (string.Equals(type, "percent", StringComparison.OrdinalIgnoreCase)) {
                double percent = TryParseThresholdNumber(threshold.Value, out double number) ? Math.Max(0D, Math.Min(100D, number)) : 0D;
                return min + ((max - min) * percent / 100D);
            }

            if (string.Equals(type, "percentile", StringComparison.OrdinalIgnoreCase)) {
                double percentile = TryParseThresholdNumber(threshold.Value, out double number) ? Math.Max(0D, Math.Min(100D, number)) : 0D;
                return sortedValues == null
                    ? CalculatePercentile(values, percentile)
                    : CalculatePercentileSorted(sortedValues, percentile);
            }

            return fallback;
        }

        internal static double CalculatePercentile(IReadOnlyList<double> values, double percentile) {
            if (values.Count == 0) {
                return 0D;
            }

            List<double> sorted = values.OrderBy(value => value).ToList();
            return CalculatePercentileSorted(sorted, percentile);
        }

        private static double CalculatePercentileSorted(IReadOnlyList<double> sorted, double percentile) {
            if (sorted.Count == 1) {
                return sorted[0];
            }

            double position = (sorted.Count - 1) * percentile / 100D;
            int lower = (int)Math.Floor(position);
            int upper = (int)Math.Ceiling(position);
            if (lower == upper) {
                return sorted[lower];
            }

            double ratio = position - lower;
            return sorted[lower] + ((sorted[upper] - sorted[lower]) * ratio);
        }

        internal static bool TryParseThresholdNumber(string? value, out double number) {
            number = 0D;
            return !string.IsNullOrWhiteSpace(value) &&
                double.TryParse(value!.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) &&
                !double.IsNaN(number) &&
                !double.IsInfinity(number);
        }

        private static string ToRgbHex(byte red, byte green, byte blue) =>
            red.ToString("X2", CultureInfo.InvariantCulture) +
            green.ToString("X2", CultureInfo.InvariantCulture) +
            blue.ToString("X2", CultureInfo.InvariantCulture);

        private static byte InterpolateByte(byte start, byte end, double ratio) {
            return (byte)Math.Max(0, Math.Min(255, (int)Math.Round(start + ((end - start) * ratio), MidpointRounding.AwayFromZero)));
        }

        internal sealed class ColorScaleEvaluator {
            private readonly IReadOnlyList<ColorStop> _stops;

            internal ColorScaleEvaluator(IReadOnlyList<ColorStop> stops) {
                _stops = stops;
            }

            internal string GetRgbHex(double value) {
                if (value <= _stops[0].Value) {
                    return ToRgbHex(_stops[0].Red, _stops[0].Green, _stops[0].Blue);
                }

                for (int i = 0; i < _stops.Count - 1; i++) {
                    ColorStop start = _stops[i];
                    ColorStop end = _stops[i + 1];
                    if (value > end.Value) {
                        continue;
                    }

                    double ratio = end.Value <= start.Value ? 1D : Math.Max(0D, Math.Min(1D, (value - start.Value) / (end.Value - start.Value)));
                    return InterpolateRgbHex(start.Red, start.Green, start.Blue, end.Red, end.Green, end.Blue, ratio);
                }

                ColorStop last = _stops[_stops.Count - 1];
                return ToRgbHex(last.Red, last.Green, last.Blue);
            }
        }

        internal readonly struct ColorStop {
            internal ColorStop(double value, byte red, byte green, byte blue) {
                Value = value;
                Red = red;
                Green = green;
                Blue = blue;
            }

            internal double Value { get; }

            internal byte Red { get; }

            internal byte Green { get; }

            internal byte Blue { get; }
        }
    }
}
