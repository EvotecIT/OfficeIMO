using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelConditionalVisualEvaluator {
        private static List<ExcelVisualConditionalIcon> BuildConditionalIcons(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            IReadOnlyDictionary<string, int> firstStoppingPriorityByCell,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var icons = new List<ExcelVisualConditionalIcon>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase))
                .OrderBy(rule => NormalizePriority(rule.Priority))) {
                if (!TryGetIconCount(rule, out int iconCount) || !TryResolveIconSetFamily(rule, iconCount, out ExcelConditionalIconFamily family)) {
                    continue;
                }

                int priority = NormalizePriority(rule.Priority);
                List<ConditionalNumericCell> candidates = GetNumericCandidates(sheet, cells, rule.Range)
                    .Where(candidate => !WasStoppedBeforePriority(firstStoppingPriorityByCell,
                        Key(candidate.Cell.Row, candidate.Cell.Column), priority))
                    .ToList();
                if (candidates.Count == 0) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported,
                        "Conditional formatting icon set could not be rendered because no numeric cells were found in the formatted range.",
                        sheet.Name + "!" + rule.Range));
                    continue;
                }

                IReadOnlyList<double> values = GetRuleNumericValues(sheet, rule.Range);
                if (values.Count == 0) {
                    values = candidates.Select(candidate => candidate.Value).ToArray();
                }

                double min = values.Min();
                double max = values.Max();
                foreach (ConditionalNumericCell candidate in candidates) {
                    int index = ResolveIconIndex(candidate.Value, values, min, max, iconCount, rule.IconSetThresholds);
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
            TryGetIconCount(rule, out int count) && count >= 3 && count <= 5 && TryResolveIconSetFamily(rule, count, out _);

        private static bool TryResolveIconSetFamily(ExcelConditionalFormattingInfo rule, int iconCount, out ExcelConditionalIconFamily family) {
            string name = rule.IconSet ?? string.Empty;
            if (name.IndexOf("Arrow", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Arrows;
                return true;
            }

            if (name.IndexOf("Rating", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Ratings;
                return true;
            }

            if (name.IndexOf("Quarters", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Quarters;
                return true;
            }

            if (name.IndexOf("Traffic", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Circles;
                return true;
            }

            if (name.IndexOf("Flag", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Flags;
                return true;
            }

            if (name.IndexOf("Signs", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Symbols", StringComparison.OrdinalIgnoreCase) >= 0) {
                family = ExcelConditionalIconFamily.Symbols;
                return true;
            }

            family = ExcelConditionalIconFamily.Symbols;
            return string.IsNullOrWhiteSpace(name);
        }

        private static bool TryGetIconCount(ExcelConditionalFormattingInfo rule, out int count) {
            string name = rule.IconSet ?? string.Empty;
            if (name.StartsWith("Four", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith("4", StringComparison.OrdinalIgnoreCase)) {
                count = 4;
                return true;
            }

            if (name.StartsWith("Five", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith("5", StringComparison.OrdinalIgnoreCase)) {
                count = 5;
                return true;
            }

            count = 3;
            return true;
        }

        private static int ResolveIconIndex(double value, IReadOnlyList<double> values, double min, double max, int iconCount, IReadOnlyList<ExcelConditionalIconSetThreshold> thresholds) {
            int resolved = 0;
            for (int i = 1; i < iconCount; i++) {
                double threshold = ResolveIconThreshold(values, min, max, iconCount, thresholds, i);
                ExcelConditionalIconSetThreshold? thresholdInfo = thresholds.Count == iconCount && i < thresholds.Count
                    ? thresholds[i]
                    : null;
                bool matches = thresholdInfo?.GreaterThanOrEqual == false
                    ? value > threshold
                    : value >= threshold;
                if (matches) {
                    resolved = i;
                }
            }

            return Math.Max(0, Math.Min(iconCount - 1, resolved));
        }

        private static double ResolveIconThreshold(IReadOnlyList<double> values, double min, double max, int iconCount, IReadOnlyList<ExcelConditionalIconSetThreshold> thresholds, int index) {
            if (thresholds.Count == iconCount && index < thresholds.Count) {
                ExcelConditionalIconSetThreshold threshold = thresholds[index];
                if (string.Equals(threshold.Type, "num", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(threshold.Type, "Number", StringComparison.OrdinalIgnoreCase)) {
                    if (ExcelConditionalFormatThresholds.TryParseThresholdNumber(threshold.Value, out double numeric)) {
                        return numeric;
                    }
                }

                if ((string.Equals(threshold.Type, "percent", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(threshold.Type, "Percent", StringComparison.OrdinalIgnoreCase)) &&
                    ExcelConditionalFormatThresholds.TryParseThresholdNumber(threshold.Value, out double percent)) {
                    return min + ((max - min) * Math.Max(0D, Math.Min(100D, percent)) / 100D);
                }

                if (string.Equals(threshold.Type, "percentile", StringComparison.OrdinalIgnoreCase) &&
                    ExcelConditionalFormatThresholds.TryParseThresholdNumber(threshold.Value, out double percentile)) {
                    return ExcelConditionalFormatThresholds.CalculatePercentile(values, Math.Max(0D, Math.Min(100D, percentile)));
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
                if (iconCount >= 5) {
                    return normalized == 0
                        ? ExcelConditionalIconKind.RedDownArrow
                        : normalized == 1
                            ? ExcelConditionalIconKind.YellowDownArrow
                            : normalized == 2
                                ? ExcelConditionalIconKind.YellowSideArrow
                                : normalized == 3
                                    ? ExcelConditionalIconKind.YellowUpArrow
                                    : ExcelConditionalIconKind.GreenUpArrow;
                }

                if (iconCount == 4) {
                    return normalized == 0
                        ? ExcelConditionalIconKind.RedDownArrow
                        : normalized == 1
                            ? ExcelConditionalIconKind.YellowDownArrow
                            : normalized == 2
                                ? ExcelConditionalIconKind.YellowSideArrow
                                : ExcelConditionalIconKind.GreenUpArrow;
                }

                return normalized == 0
                    ? ExcelConditionalIconKind.RedDownArrow
                    : normalized == 1
                        ? ExcelConditionalIconKind.YellowSideArrow
                        : ExcelConditionalIconKind.GreenUpArrow;
            }

            if (family == ExcelConditionalIconFamily.Circles) {
                if (iconCount >= 5) {
                    return normalized == 0
                        ? ExcelConditionalIconKind.RedCircle
                        : normalized == 1
                            ? ExcelConditionalIconKind.OrangeCircle
                            : normalized == 2
                                ? ExcelConditionalIconKind.YellowCircle
                                : normalized == 3
                                    ? ExcelConditionalIconKind.LightGreenCircle
                                    : ExcelConditionalIconKind.GreenCircle;
                }

                if (iconCount == 4) {
                    return normalized == 0
                        ? ExcelConditionalIconKind.RedCircle
                        : normalized == 1
                            ? ExcelConditionalIconKind.OrangeCircle
                            : normalized == 2
                                ? ExcelConditionalIconKind.YellowCircle
                                : ExcelConditionalIconKind.GreenCircle;
                }

                return normalized == 0
                    ? ExcelConditionalIconKind.RedCircle
                    : normalized == 1
                        ? ExcelConditionalIconKind.YellowCircle
                        : ExcelConditionalIconKind.GreenCircle;
            }

            if (family == ExcelConditionalIconFamily.Ratings) {
                if (iconCount >= 5) {
                    return normalized == 0
                        ? ExcelConditionalIconKind.RatingOne
                        : normalized == 1
                            ? ExcelConditionalIconKind.RatingTwo
                            : normalized == 2
                                ? ExcelConditionalIconKind.RatingThree
                                : normalized == 3
                                    ? ExcelConditionalIconKind.RatingFour
                                    : ExcelConditionalIconKind.RatingFive;
                }

                return normalized == 0
                    ? ExcelConditionalIconKind.RatingOne
                    : normalized == 1
                        ? ExcelConditionalIconKind.RatingTwo
                        : normalized == 2
                            ? ExcelConditionalIconKind.RatingThree
                            : ExcelConditionalIconKind.RatingFour;
            }

            if (family == ExcelConditionalIconFamily.Quarters) {
                return normalized == 0
                    ? ExcelConditionalIconKind.QuarterEmpty
                    : normalized == 1
                        ? ExcelConditionalIconKind.QuarterOne
                        : normalized == 2
                            ? ExcelConditionalIconKind.QuarterTwo
                            : normalized == 3
                                ? ExcelConditionalIconKind.QuarterThree
                                : ExcelConditionalIconKind.QuarterFull;
            }

            if (family == ExcelConditionalIconFamily.Flags) {
                return normalized == 0
                    ? ExcelConditionalIconKind.RedFlag
                    : normalized == 1
                        ? ExcelConditionalIconKind.YellowFlag
                        : ExcelConditionalIconKind.GreenFlag;
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
            Circles,
            Ratings,
            Quarters,
            Flags
        }
    }
}
