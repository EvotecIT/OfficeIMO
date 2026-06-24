using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelConditionalVisualEvaluator {
        private static List<ExcelVisualConditionalIcon> BuildConditionalIcons(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var icons = new List<ExcelVisualConditionalIcon>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase))
                .OrderBy(rule => NormalizePriority(rule.Priority))) {
                if (!TryResolveIconSetFamily(rule, out ExcelConditionalIconFamily family) || !TryGetIconCount(rule, out int iconCount)) {
                    continue;
                }

                List<ConditionalNumericCell> candidates = GetNumericCandidates(cells, rule.Range);
                if (candidates.Count == 0) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported,
                        "Conditional formatting icon set could not be rendered because no numeric cells were found in the formatted range.",
                        sheet.Name + "!" + rule.Range));
                    continue;
                }

                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (ConditionalNumericCell candidate in candidates) {
                    int index = ResolveIconIndex(candidate.Value, min, max, iconCount, rule.IconSetThresholds);
                    if (rule.IconSetReverse) {
                        index = (iconCount - 1) - index;
                    }

                    icons.Add(new ExcelVisualConditionalIcon(
                        candidate.Cell.Row,
                        candidate.Cell.Column,
                        candidate.Cell.X,
                        candidate.Cell.Y,
                        candidate.Cell.Width,
                        candidate.Cell.Height,
                        MapIconKind(family, index, iconCount),
                        rule.IconSetShowValue));
                }
            }

            return icons;
        }

        private static bool CanRenderIconSet(ExcelConditionalFormattingInfo rule) =>
            TryResolveIconSetFamily(rule, out _) && TryGetIconCount(rule, out int count) && count == 3;

        private static bool TryResolveIconSetFamily(ExcelConditionalFormattingInfo rule, out ExcelConditionalIconFamily family) {
            string name = rule.IconSet ?? string.Empty;
            if (name.IndexOf("Traffic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Signs", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Symbols", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Symbols;
                return true;
            }

            if (name.IndexOf("Arrow", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Arrows;
                return true;
            }

            if (name.IndexOf("Rating", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Quarters", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Circles;
                return true;
            }

            family = ExcelConditionalIconFamily.Symbols;
            return string.IsNullOrWhiteSpace(name);
        }

        private static bool TryGetIconCount(ExcelConditionalFormattingInfo rule, out int count) {
            string name = rule.IconSet ?? string.Empty;
            if (name.StartsWith("Four", StringComparison.OrdinalIgnoreCase)) {
                count = 4;
                return false;
            }

            if (name.StartsWith("Five", StringComparison.OrdinalIgnoreCase)) {
                count = 5;
                return false;
            }

            count = 3;
            return true;
        }

        private static int ResolveIconIndex(double value, double min, double max, int iconCount, IReadOnlyList<ExcelConditionalIconSetThreshold> thresholds) {
            int resolved = 0;
            for (int i = 1; i < iconCount; i++) {
                double threshold = ResolveIconThreshold(min, max, iconCount, thresholds, i);
                if (value >= threshold) {
                    resolved = i;
                }
            }

            return Math.Max(0, Math.Min(iconCount - 1, resolved));
        }

        private static double ResolveIconThreshold(double min, double max, int iconCount, IReadOnlyList<ExcelConditionalIconSetThreshold> thresholds, int index) {
            if (thresholds.Count == iconCount && index < thresholds.Count) {
                ExcelConditionalIconSetThreshold threshold = thresholds[index];
                if (string.Equals(threshold.Type, "num", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(threshold.Type, "Number", StringComparison.OrdinalIgnoreCase)) {
                    if (double.TryParse(threshold.Value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double numeric)) {
                        return numeric;
                    }
                }

                if ((string.Equals(threshold.Type, "percent", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(threshold.Type, "Percent", StringComparison.OrdinalIgnoreCase)) &&
                    double.TryParse(threshold.Value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double percent)) {
                    return min + ((max - min) * Math.Max(0D, Math.Min(100D, percent)) / 100D);
                }
            }

            double fallbackPercent = iconCount == 3
                ? index == 1 ? 33D : 67D
                : (100D / iconCount) * index;
            return min + ((max - min) * fallbackPercent / 100D);
        }

        private static ExcelConditionalIconKind MapIconKind(ExcelConditionalIconFamily family, int index, int iconCount) {
            int normalized = Math.Max(0, Math.Min(iconCount - 1, index));
            if (family == ExcelConditionalIconFamily.Arrows) {
                return normalized == 0
                    ? ExcelConditionalIconKind.RedDownArrow
                    : normalized == 1
                        ? ExcelConditionalIconKind.YellowSideArrow
                        : ExcelConditionalIconKind.GreenUpArrow;
            }

            if (family == ExcelConditionalIconFamily.Circles) {
                return normalized == 0
                    ? ExcelConditionalIconKind.RedCircle
                    : normalized == 1
                        ? ExcelConditionalIconKind.YellowCircle
                        : ExcelConditionalIconKind.GreenCircle;
            }

            return normalized == 0
                ? ExcelConditionalIconKind.RedCross
                : normalized == 1
                    ? ExcelConditionalIconKind.YellowExclamation
                    : ExcelConditionalIconKind.GreenCheck;
        }

        private enum ExcelConditionalIconFamily {
            Symbols,
            Arrows,
            Circles
        }
    }
}
