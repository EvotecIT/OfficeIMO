using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelConditionalVisualEvaluator {
        private const int MaxConditionalReferenceCells = 100_000;

        internal static ExcelConditionalVisualState Evaluate(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            string range,
            DateTime conditionalFormattingDate,
            List<OfficeImageExportDiagnostic> diagnostics) {
            IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules(range);
            if (rules.Count == 0 || cells.Count == 0) {
                return ExcelConditionalVisualState.Empty;
            }

            rules = ExcludeOversizedConditionalRules(sheet, rules, diagnostics);
            if (rules.Count == 0) {
                return ExcelConditionalVisualState.Empty;
            }

            ReportUnsupportedConditionalRules(sheet, cells, rules, conditionalFormattingDate, diagnostics);

            IReadOnlyDictionary<string, int> firstStoppingPriorityByCell =
                BuildFirstStoppingPriorityByCell(sheet, cells, rules, conditionalFormattingDate);
            var stoppedCells = new HashSet<string>(StringComparer.Ordinal);
            var cellFormats = BuildConditionalCellFormats(sheet, cells, rules, conditionalFormattingDate, stoppedCells);
            var dataBars = BuildConditionalDataBars(sheet, cells, rules, firstStoppingPriorityByCell);
            var icons = BuildConditionalIcons(sheet, cells, rules, firstStoppingPriorityByCell, diagnostics);
            return cellFormats.Count == 0 && dataBars.Count == 0 && icons.Count == 0
                ? ExcelConditionalVisualState.Empty
                : new ExcelConditionalVisualState(cellFormats, dataBars, icons);
        }

        private static IReadOnlyList<ExcelConditionalFormattingInfo> ExcludeOversizedConditionalRules(
            ExcelSheet sheet,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var retained = new List<ExcelConditionalFormattingInfo>(rules.Count);
            foreach (ExcelConditionalFormattingInfo rule in rules) {
                if (!EnumerateReferenceCells(rule.Range, MaxConditionalReferenceCells + 1)
                    .Skip(MaxConditionalReferenceCells)
                    .Any()) {
                    retained.Add(rule);
                    continue;
                }

                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded,
                    $"Conditional formatting rule was omitted because its reference exceeds the {MaxConditionalReferenceCells.ToString(CultureInfo.InvariantCulture)}-cell image-export limit.",
                    sheet.Name + "!" + rule.Range));
            }

            return retained;
        }

        private static List<ExcelVisualConditionalDataBar> BuildConditionalDataBars(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            IReadOnlyDictionary<string, int> firstStoppingPriorityByCell) {
            var dataBars = new List<ExcelVisualConditionalDataBar>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "DataBar", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.DataBarColor))
                .OrderBy(rule => NormalizePriority(rule.Priority))) {
                if (!TryNormalizeArgb(rule.DataBarColor, out string? colorArgb) || colorArgb == null) {
                    continue;
                }

                int priority = NormalizePriority(rule.Priority);
                List<ConditionalNumericCell> candidates = GetNumericCandidates(sheet, cells, rule.Range)
                    .Where(candidate => !WasStoppedBeforePriority(firstStoppingPriorityByCell,
                        Key(candidate.Cell.Row, candidate.Cell.Column), priority))
                    .ToList();
                if (candidates.Count == 0) {
                    continue;
                }

                IReadOnlyList<double> values = GetRuleNumericValues(sheet, rule.Range);
                if (values.Count == 0) {
                    values = candidates.Select(candidate => candidate.Value).ToArray();
                }

                (double min, double max) = ExcelConditionalFormatThresholds.ResolveDataBarRange(values, rule.DataBarThresholds);
                foreach (ConditionalNumericCell candidate in candidates) {
                    (double startRatio, double ratio) = ExcelConditionalFormatThresholds.GetDataBarGeometry(candidate.Value, min, max);
                    dataBars.Add(new ExcelVisualConditionalDataBar(
                        candidate.Cell.Row,
                        candidate.Cell.Column,
                        candidate.Cell.X,
                        candidate.Cell.Y,
                        candidate.Cell.Width,
                        candidate.Cell.Height,
                        colorArgb,
                        startRatio,
                        ratio,
                        rule.DataBarShowValue));
                }
            }

            return dataBars;
        }

        private static IReadOnlyDictionary<string, int> BuildFirstStoppingPriorityByCell(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            DateTime conditionalFormattingDate) {
            var stoppedCells = new HashSet<string>(StringComparer.Ordinal);
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (IGrouping<int, ExcelConditionalFormattingInfo> priorityGroup in rules
                .OrderBy(rule => NormalizePriority(rule.Priority))
                .GroupBy(rule => NormalizePriority(rule.Priority))) {
                foreach (ExcelConditionalFormattingInfo stopRule in priorityGroup.Where(rule => rule.StopIfTrue)) {
                    Dictionary<string, ExcelConditionalCellFormat> newlyStopped = BuildConditionalCellFormats(
                        sheet,
                        cells,
                        new[] { stopRule },
                        conditionalFormattingDate,
                        stoppedCells);
                    foreach (string key in newlyStopped.Keys) {
                        if (!result.ContainsKey(key)) {
                            result.Add(key, priorityGroup.Key);
                        }
                    }
                }
            }
            return result;
        }

        private static bool WasStoppedBeforePriority(
            IReadOnlyDictionary<string, int> firstStoppingPriorityByCell,
            string key,
            int priority) =>
            firstStoppingPriorityByCell.TryGetValue(key, out int stoppingPriority) && stoppingPriority < priority;

        private static void ReportUnsupportedConditionalRules(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            DateTime conditionalFormattingDate,
            List<OfficeImageExportDiagnostic> diagnostics) {
            foreach (ExcelConditionalFormattingInfo rule in rules) {
                string source = sheet.Name + "!" + rule.Range;
                if (string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase)) {
                    if (CanRenderIconSet(rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Info,
                            ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation,
                            "Conditional formatting icon set is rendered as a deterministic dependency-free approximation; Excel-specific icon artwork and threshold semantics may differ.",
                            source));
                    } else {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported,
                            "Conditional formatting icon set could not be rendered because this icon-set family is not supported yet.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "ColorScale", StringComparison.OrdinalIgnoreCase)) {
                    if (rule.ColorScaleColors.Count < 2 ||
                        !TryNormalizeArgb(rule.ColorScaleColors[0], out _) ||
                        !TryNormalizeArgb(rule.ColorScaleColors[rule.ColorScaleColors.Count - 1], out _)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalColorScaleUnsupported,
                            "Conditional formatting color scale could not be rendered because its color stops are missing or unsupported.",
                            source));
                    }

                    if (ExcelConditionalFormatThresholds.HasUnsupportedFormulaThresholds(rule.ColorScaleThresholds)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalFormulaThresholdApproximation,
                            "Conditional formatting color-scale formula thresholds are not evaluated by image export; threshold fallback positions were used.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "DataBar", StringComparison.OrdinalIgnoreCase)) {
                    if (!TryNormalizeArgb(rule.DataBarColor, out _)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalDataBarUnsupported,
                            "Conditional formatting data bar could not be rendered because its fill color is missing or unsupported.",
                            source));
                    }

                    if (ExcelConditionalFormatThresholds.HasUnsupportedFormulaThresholds(rule.DataBarThresholds)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalFormulaThresholdApproximation,
                            "Conditional formatting data-bar formula thresholds are not evaluated by image export; threshold fallback positions were used.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "CellIs", StringComparison.OrdinalIgnoreCase)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateAnyCellIsRule(sheet, cells, rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalCellIsUnsupported,
                            "Conditional formatting cell-is rule was not rendered because only bounded numeric comparisons are supported.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "Expression", StringComparison.OrdinalIgnoreCase)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateAnyExpressionRule(sheet, cells, rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalFormulaUnsupported,
                            "Conditional formatting formula rule was not rendered because only simple numeric comparison expressions are supported.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "Top10", StringComparison.OrdinalIgnoreCase)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateTopBottomRule(sheet, cells, rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalTopBottomUnsupported,
                            "Conditional formatting top/bottom rule was not rendered because it has no valid numeric candidates or rank.",
                            source));
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "DuplicateValues", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rule.Type, "UniqueValues", StringComparison.OrdinalIgnoreCase)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    continue;
                }

                if (string.Equals(rule.Type, "AboveAverage", StringComparison.OrdinalIgnoreCase)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateAboveAverageRule(sheet, cells, rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalAboveAverageUnsupported,
                            "Conditional formatting above/below-average rule was not rendered because only numeric average rules without standard-deviation thresholds are supported.",
                            source));
                    }

                    continue;
                }

                if (IsTextRule(rule)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateTextRule(rule)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalTextRuleUnsupported,
                            "Conditional formatting text rule was not rendered because its comparison text is missing.",
                            source));
                    }

                    continue;
                }

                if (IsTimePeriodRule(rule)) {
                    if (ReportUnsupportedDifferentialFormat(rule, diagnostics, source)) {
                        continue;
                    }

                    if (HasSupportedDifferentialFormat(rule) &&
                        !CanEvaluateTimePeriodRule(sheet, cells, rule, conditionalFormattingDate)) {
                        diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ConditionalTimePeriodUnsupported,
                            "Conditional formatting time-period rule was not rendered because its time period is missing, unsupported, or no valid date cells were found.",
                            source));
                    }

                    continue;
                }

                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported,
                    "Conditional formatting rule type is not rendered by Excel image export yet.",
                    source));
            }
        }

        private static bool ReportUnsupportedDifferentialFormat(
            ExcelConditionalFormattingInfo rule,
            List<OfficeImageExportDiagnostic> diagnostics,
            string source) {
            if (!rule.DifferentialFormatId.HasValue || HasSupportedDifferentialFormat(rule)) {
                return false;
            }

            diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.ConditionalDifferentialFormatUnsupported,
                "Conditional formatting differential format does not contain a supported solid fill, font effect, or border; number-format and other differential effects are not rendered yet.",
                source));
            return true;
        }

        private static Dictionary<string, ExcelConditionalCellFormat> BuildConditionalCellFormats(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelConditionalFormattingInfo> rules,
            DateTime conditionalFormattingDate,
            HashSet<string> stoppedCells) {
            var formats = new Dictionary<string, ExcelConditionalCellFormat>(StringComparer.Ordinal);
            foreach (ExcelConditionalFormattingInfo rule in rules.OrderBy(rule => NormalizePriority(rule.Priority))) {
                if (string.Equals(rule.Type, "ColorScale", StringComparison.OrdinalIgnoreCase)) {
                    ApplyColorScaleFill(sheet, cells, rule, formats, stoppedCells);
                    continue;
                }

                if (!string.Equals(rule.Type, "CellIs", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.Type, "Expression", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.Type, "Top10", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.Type, "DuplicateValues", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.Type, "UniqueValues", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rule.Type, "AboveAverage", StringComparison.OrdinalIgnoreCase) &&
                    !IsTextRule(rule) &&
                    !IsTimePeriodRule(rule)) {
                    continue;
                }

                if (string.Equals(rule.Type, "Top10", StringComparison.OrdinalIgnoreCase)) {
                    ApplyTopBottomFormat(sheet, cells, rule, formats, stoppedCells);
                    continue;
                }

                if (string.Equals(rule.Type, "DuplicateValues", StringComparison.OrdinalIgnoreCase)) {
                    ApplyDistinctValueFormat(sheet, cells, rule, formats, stoppedCells, selectDuplicates: true);
                    continue;
                }

                if (string.Equals(rule.Type, "UniqueValues", StringComparison.OrdinalIgnoreCase)) {
                    ApplyDistinctValueFormat(sheet, cells, rule, formats, stoppedCells, selectDuplicates: false);
                    continue;
                }

                if (string.Equals(rule.Type, "AboveAverage", StringComparison.OrdinalIgnoreCase)) {
                    ApplyAboveAverageFormat(sheet, cells, rule, formats, stoppedCells);
                    continue;
                }

                if (IsTextRule(rule)) {
                    ApplyTextRuleFormat(cells, rule, formats, stoppedCells);
                    continue;
                }

                if (IsTimePeriodRule(rule)) {
                    ApplyTimePeriodFormat(sheet, cells, rule, conditionalFormattingDate, formats, stoppedCells);
                    continue;
                }

                foreach (ExcelVisualCell cell in cells) {
                    string key = Key(cell.Row, cell.Column);
                    if (cell.CoveredByMerge || stoppedCells.Contains(key) || !IsCellInReferenceList(cell.Row, cell.Column, rule.Range)) {
                        continue;
                    }

                    if (RuleMatchesCell(sheet, cell, rule)) {
                        ApplyDifferentialFormat(rule, key, formats);

                        if (rule.StopIfTrue) {
                            stoppedCells.Add(key);
                        }
                    }
                }
            }

            return formats;
        }

        private static void ApplyColorScaleFill(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            Dictionary<string, ExcelConditionalCellFormat> formats,
            HashSet<string> stoppedCells) {
            if (rule.ColorScaleColors.Count < 2 ||
                !ExcelConditionalFormatThresholds.TryGetRgb(rule.ColorScaleColors[0], out _, out _, out _) ||
                !ExcelConditionalFormatThresholds.TryGetRgb(rule.ColorScaleColors[rule.ColorScaleColors.Count - 1], out _, out _, out _)) {
                return;
            }

            List<ConditionalNumericCell> candidates = GetNumericCandidates(sheet, cells, rule.Range)
                .Where(candidate => !stoppedCells.Contains(Key(candidate.Cell.Row, candidate.Cell.Column)))
                .ToList();
            if (candidates.Count == 0) {
                return;
            }

            IReadOnlyList<double> values = GetRuleNumericValues(sheet, rule.Range, Array.Empty<string>());
            if (values.Count == 0) {
                values = candidates.Select(candidate => candidate.Value).ToArray();
            }

            foreach (ConditionalNumericCell candidate in candidates) {
                string key = Key(candidate.Cell.Row, candidate.Cell.Column);
                if (formats.TryGetValue(key, out ExcelConditionalCellFormat? existing) &&
                    !string.IsNullOrWhiteSpace(existing.FillColorArgb)) {
                    continue;
                }

                if (ExcelConditionalFormatThresholds.TryGetColorScaleRgb(values, rule.ColorScaleColors, rule.ColorScaleThresholds, candidate.Value, out string rgbHex)) {
                    ApplyFillFormat(key, "FF" + rgbHex, formats);
                }
            }
        }

        private static bool RuleMatchesCell(ExcelSheet sheet, ExcelVisualCell cell, ExcelConditionalFormattingInfo rule) {
            if (string.Equals(rule.Type, "CellIs", StringComparison.OrdinalIgnoreCase)) {
                return MatchesCellIsRule(sheet, cell, rule);
            }

            return string.Equals(rule.Type, "Expression", StringComparison.OrdinalIgnoreCase) &&
                rule.Formulas.Count > 0 &&
                TryEvaluateComparisonExpression(sheet, cell, rule.Range, rule.Formulas[0], out bool result) &&
                result;
        }

        private static bool CanEvaluateTopBottomRule(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, ExcelConditionalFormattingInfo rule) {
            if (!rule.TopBottomRank.HasValue || rule.TopBottomRank.Value == 0U) {
                return false;
            }

            int count = GetRuleNumericValues(sheet, rule.Range, Array.Empty<string>()).Count;
            if (count == 0) {
                count = GetNumericCandidates(sheet, cells, rule.Range).Count;
            }

            return CalculateTopBottomSelectionCount(rule, count) > 0;
        }

        private static void ApplyTopBottomFormat(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            Dictionary<string, ExcelConditionalCellFormat> formats,
            HashSet<string> stoppedCells) {
            if (!rule.TopBottomRank.HasValue ||
                rule.TopBottomRank.Value == 0U) {
                return;
            }

            List<ConditionalNumericCell> candidates = GetNumericCandidates(sheet, cells, rule.Range)
                .Where(candidate => !stoppedCells.Contains(Key(candidate.Cell.Row, candidate.Cell.Column)))
                .ToList();
            if (candidates.Count == 0) {
                return;
            }

            IReadOnlyList<double> values = GetRuleNumericValues(sheet, rule.Range, Array.Empty<string>());
            if (values.Count == 0) {
                values = candidates.Select(candidate => candidate.Value).ToArray();
            }

            int rank = CalculateTopBottomSelectionCount(rule, values.Count);
            if (rank == 0) {
                return;
            }

            List<double> orderedValues = values
                .OrderBy(value => rule.TopBottomBottom ? value : -value)
                .ToList();
            double cutoff = orderedValues[rank - 1];
            foreach (ConditionalNumericCell candidate in candidates) {
                bool selected = rule.TopBottomBottom
                    ? candidate.Value <= cutoff
                    : candidate.Value >= cutoff;
                if (!selected) {
                    continue;
                }

                string key = Key(candidate.Cell.Row, candidate.Cell.Column);
                ApplyDifferentialFormat(rule, key, formats);

                if (rule.StopIfTrue) {
                    stoppedCells.Add(key);
                }
            }
        }

        private static int CalculateTopBottomSelectionCount(ExcelConditionalFormattingInfo rule, int candidateCount) {
            if (candidateCount <= 0 || !rule.TopBottomRank.HasValue || rule.TopBottomRank.Value == 0U) {
                return 0;
            }

            if (!rule.TopBottomPercent) {
                return (int)Math.Min(rule.TopBottomRank.Value, (uint)candidateCount);
            }

            double percent = Math.Min(100D, rule.TopBottomRank.Value);
            int count = (int)Math.Ceiling(candidateCount * (percent / 100D));
            return Math.Max(1, Math.Min(candidateCount, count));
        }

        private static bool CanEvaluateAboveAverageRule(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, ExcelConditionalFormattingInfo rule) {
            if (rule.AboveAverageStdDev.HasValue) {
                return false;
            }

            int count = GetRuleNumericValues(sheet, rule.Range, Array.Empty<string>()).Count;
            if (count == 0) {
                count = GetNumericCandidates(sheet, cells, rule.Range).Count;
            }

            return count > 0;
        }

        private static void ApplyAboveAverageFormat(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            Dictionary<string, ExcelConditionalCellFormat> formats,
            HashSet<string> stoppedCells) {
            if (rule.AboveAverageStdDev.HasValue) {
                return;
            }

            List<ConditionalNumericCell> candidates = GetNumericCandidates(sheet, cells, rule.Range);
            if (candidates.Count == 0) {
                return;
            }

            IReadOnlyList<double> values = GetRuleNumericValues(sheet, rule.Range, Array.Empty<string>());
            if (values.Count == 0) {
                values = candidates.Select(candidate => candidate.Value).ToArray();
            }

            double average = values.Average();
            foreach (ConditionalNumericCell candidate in candidates) {
                string key = Key(candidate.Cell.Row, candidate.Cell.Column);
                if (stoppedCells.Contains(key)) {
                    continue;
                }

                bool selected = rule.AboveAverageAbove
                    ? rule.AboveAverageEqual ? candidate.Value >= average : candidate.Value > average
                    : rule.AboveAverageEqual ? candidate.Value <= average : candidate.Value < average;
                if (!selected) {
                    continue;
                }

                ApplyDifferentialFormat(rule, key, formats);

                if (rule.StopIfTrue) {
                    stoppedCells.Add(key);
                }
            }
        }

        private static void ApplyDistinctValueFormat(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            Dictionary<string, ExcelConditionalCellFormat> formats,
            HashSet<string> stoppedCells,
            bool selectDuplicates) {
            var candidates = new List<(ExcelVisualCell Cell, string Value)>();
            foreach (ExcelVisualCell cell in GetRuleCells(cells, rule.Range)) {
                string key = Key(cell.Row, cell.Column);
                if (stoppedCells.Contains(key) || string.IsNullOrWhiteSpace(cell.Text)) {
                    continue;
                }

                string value = TryGetCellTextValue(sheet, cell.Row, cell.Column, out string rawValue)
                    ? rawValue
                    : cell.Text.Trim();
                candidates.Add((cell, value));
            }

            if (candidates.Count == 0) {
                return;
            }

            IReadOnlyList<string> values = GetRuleTextValues(sheet, rule.Range);
            if (values.Count == 0) {
                values = candidates.Select(candidate => candidate.Value).ToArray();
            }

            var selectedValues = new HashSet<string>(
                values
                    .GroupBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .Where(group => selectDuplicates ? group.Count() > 1 : group.Count() == 1)
                    .Select(group => group.Key),
                StringComparer.OrdinalIgnoreCase);
            if (selectedValues.Count == 0) {
                return;
            }

            foreach ((ExcelVisualCell cell, string value) in candidates) {
                if (!selectedValues.Contains(value)) {
                    continue;
                }

                string key = Key(cell.Row, cell.Column);
                ApplyDifferentialFormat(rule, key, formats);

                if (rule.StopIfTrue) {
                    stoppedCells.Add(key);
                }
            }
        }

        private static bool MatchesCellIsRule(ExcelSheet sheet, ExcelVisualCell cell, ExcelConditionalFormattingInfo rule) {
            if (rule.Formulas.Count == 0 || string.IsNullOrWhiteSpace(rule.Operator)) {
                return false;
            }

            if (!TryGetCellNumericValue(sheet, cell, out double cellValue) ||
                !TryResolveNumericOperand(sheet, cell, rule.Range, rule.Formulas[0], out double firstValue)) {
                return false;
            }

            if (string.Equals(rule.Operator, "Between", StringComparison.OrdinalIgnoreCase)) {
                return rule.Formulas.Count > 1 &&
                    TryResolveNumericOperand(sheet, cell, rule.Range, rule.Formulas[1], out double secondValue) &&
                    cellValue >= Math.Min(firstValue, secondValue) &&
                    cellValue <= Math.Max(firstValue, secondValue);
            }

            if (string.Equals(rule.Operator, "NotBetween", StringComparison.OrdinalIgnoreCase)) {
                return rule.Formulas.Count > 1 &&
                    TryResolveNumericOperand(sheet, cell, rule.Range, rule.Formulas[1], out double secondValue) &&
                    (cellValue < Math.Min(firstValue, secondValue) || cellValue > Math.Max(firstValue, secondValue));
            }

            return Compare(cellValue, firstValue, rule.Operator!);
        }

        private static bool CanEvaluateAnyCellIsRule(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, ExcelConditionalFormattingInfo rule) {
            if (rule.Formulas.Count == 0 || string.IsNullOrWhiteSpace(rule.Operator) || !IsSupportedCellIsOperator(rule.Operator!)) {
                return false;
            }

            foreach (ExcelVisualCell cell in GetRuleCells(cells, rule.Range)) {
                if (!TryGetCellNumericValue(sheet, cell, out _) ||
                    !TryResolveNumericOperand(sheet, cell, rule.Range, rule.Formulas[0], out _)) {
                    continue;
                }

                if (string.Equals(rule.Operator, "Between", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(rule.Operator, "NotBetween", StringComparison.OrdinalIgnoreCase)) {
                    if (rule.Formulas.Count > 1 &&
                        TryResolveNumericOperand(sheet, cell, rule.Range, rule.Formulas[1], out _)) {
                        return true;
                    }

                    continue;
                }

                return true;
            }

            return false;
        }

        private static bool CanEvaluateAnyExpressionRule(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, ExcelConditionalFormattingInfo rule) {
            if (rule.Formulas.Count == 0) {
                return false;
            }

            foreach (ExcelVisualCell cell in GetRuleCells(cells, rule.Range)) {
                if (TryEvaluateComparisonExpression(sheet, cell, rule.Range, rule.Formulas[0], out _)) {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<ExcelVisualCell> GetRuleCells(IReadOnlyList<ExcelVisualCell> cells, string referenceList) {
            foreach (ExcelVisualCell cell in cells) {
                if (!cell.CoveredByMerge && IsCellInReferenceList(cell.Row, cell.Column, referenceList)) {
                    yield return cell;
                }
            }
        }

        private static bool IsSupportedCellIsOperator(string op) {
            return string.Equals(op, "Between", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "NotBetween", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "GreaterThan", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "LessThan", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "GreaterThanOrEqual", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "LessThanOrEqual", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "Equal", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(op, "NotEqual", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryEvaluateComparisonExpression(ExcelSheet sheet, ExcelVisualCell cell, string ruleRange, string formula, out bool result) {
            result = false;
            string expression = NormalizeFormula(formula);
            string[] operators = { ">=", "<=", "<>", "=", ">", "<" };
            foreach (string op in operators) {
                int index = expression.IndexOf(op, StringComparison.Ordinal);
                if (index <= 0 || index + op.Length >= expression.Length) {
                    continue;
                }

                string left = expression.Substring(0, index).Trim();
                string right = expression.Substring(index + op.Length).Trim();
                if (TryResolveNumericOperand(sheet, cell, ruleRange, left, out double leftValue) &&
                    TryResolveNumericOperand(sheet, cell, ruleRange, right, out double rightValue)) {
                    result = Compare(leftValue, rightValue, op);
                    return true;
                }
            }

            return false;
        }

        private static bool Compare(double left, double right, string op) {
            return op switch {
                "GreaterThan" or ">" => left > right,
                "LessThan" or "<" => left < right,
                "GreaterThanOrEqual" or ">=" => left >= right,
                "LessThanOrEqual" or "<=" => left <= right,
                "Equal" or "=" => Math.Abs(left - right) < 0.0000001D,
                "NotEqual" or "<>" => Math.Abs(left - right) >= 0.0000001D,
                _ => false
            };
        }

        private static bool TryResolveNumericOperand(ExcelSheet sheet, ExcelVisualCell cell, string ruleRange, string operand, out double value) {
            value = 0D;
            string normalized = NormalizeFormula(operand);
            if (double.TryParse(normalized, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value)) {
                return true;
            }

            return TryResolveCellReference(sheet, cell, ruleRange, normalized, out int row, out int column) &&
                TryGetCellNumericValue(sheet, row, column, out value);
        }

        private static bool TryResolveCellReference(ExcelSheet sheet, ExcelVisualCell cell, string ruleRange, string reference, out int row, out int column) {
            row = column = 0;
            string normalized = NormalizeCellReference(StripSheetPrefix(reference), out bool absoluteColumn, out bool absoluteRow);
            if (!A1.TryParseCellReferenceFast(normalized, out int referenceRow, out int referenceColumn)) {
                return false;
            }

            (int topRow, int leftColumn) = GetReferenceListOrigin(ruleRange);
            row = absoluteRow ? referenceRow : referenceRow + (cell.Row - topRow);
            column = absoluteColumn ? referenceColumn : referenceColumn + (cell.Column - leftColumn);
            return row >= 1 && column >= 1;
        }

        private static string NormalizeCellReference(string reference, out bool absoluteColumn, out bool absoluteRow) {
            absoluteColumn = false;
            absoluteRow = false;
            if (string.IsNullOrWhiteSpace(reference)) {
                return string.Empty;
            }

            int index = 0;
            if (reference[index] == '$') {
                absoluteColumn = true;
                index++;
            }

            while (index < reference.Length && char.IsLetter(reference[index])) {
                index++;
            }

            if (index < reference.Length && reference[index] == '$') {
                absoluteRow = true;
            }

            return reference.Replace("$", string.Empty);
        }

        private static (int Row, int Column) GetReferenceListOrigin(string referenceList) {
            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out _, out _)) {
                    return (firstRow, firstColumn);
                }

                if (A1.TryParseCellReferenceFast(token, out int row, out int column)) {
                    return (row, column);
                }
            }

            return (1, 1);
        }

        private static string NormalizeFormula(string formula) {
            string normalized = formula.Trim();
            return normalized.StartsWith("=", StringComparison.Ordinal) ? normalized.Substring(1).Trim() : normalized;
        }

        private static int NormalizePriority(int priority) => priority <= 0 ? int.MaxValue : priority;

        private static bool TryGetCellNumericValue(ExcelSheet sheet, ExcelVisualCell cell, out double value) {
            return TryGetCellNumericValue(sheet, cell.Row, cell.Column, out value) ||
                TryGetConditionalNumericValue(cell.Text, out value);
        }

        private static bool TryGetCellNumericValue(ExcelSheet sheet, int row, int column, out double value) {
            ExcelCellData data = sheet.GetCellValueSnapshot(row, column);
            if (data.Value is double doubleValue) {
                value = doubleValue;
                return true;
            }

            if (data.Value is IConvertible convertible &&
                double.TryParse(Convert.ToString(convertible, CultureInfo.InvariantCulture), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value)) {
                return true;
            }

            return TryGetConditionalNumericValue(data.CachedText, out value);
        }

        private static List<ConditionalNumericCell> GetNumericCandidates(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, string referenceList) {
            var candidates = new List<ConditionalNumericCell>();
            foreach (ExcelVisualCell cell in cells) {
                if (cell.CoveredByMerge || !IsCellInReferenceList(cell.Row, cell.Column, referenceList)) {
                    continue;
                }

                if (TryGetCellNumericValue(sheet, cell, out double value)) {
                    candidates.Add(new ConditionalNumericCell(cell, value));
                }
            }

            return candidates;
        }

        private static List<double> GetRuleNumericValues(ExcelSheet sheet, string referenceList) =>
            GetRuleNumericValues(sheet, referenceList, Array.Empty<string>());

        private static List<double> GetRuleNumericValues(ExcelSheet sheet, string referenceList, IReadOnlyCollection<string> excludedKeys) {
            var values = new List<double>();
            foreach ((int row, int column) in EnumerateReferenceCells(referenceList)) {
                if (excludedKeys.Contains(Key(row, column))) {
                    continue;
                }

                if (TryGetCellNumericValue(sheet, row, column, out double value)) {
                    values.Add(value);
                }
            }

            return values;
        }

        private static List<string> GetRuleTextValues(ExcelSheet sheet, string referenceList) {
            var values = new List<string>();
            foreach ((int row, int column) in EnumerateReferenceCells(referenceList)) {
                if (TryGetCellTextValue(sheet, row, column, out string value)) {
                    values.Add(value);
                }
            }

            return values;
        }

        private static bool TryGetCellTextValue(ExcelSheet sheet, int row, int column, out string value) {
            ExcelCellData data = sheet.GetCellValueSnapshot(row, column);
            string? text = data.CachedText;
            if (string.IsNullOrWhiteSpace(text) && data.Value != null) {
                text = Convert.ToString(data.Value, CultureInfo.InvariantCulture);
            }

            if (!string.IsNullOrWhiteSpace(text)) {
                value = text!.Trim();
                return true;
            }

            value = string.Empty;
            return false;
        }

        private static IEnumerable<(int Row, int Column)> EnumerateReferenceCells(
            string referenceList,
            int maximumCells = MaxConditionalReferenceCells) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                yield break;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            int emitted = 0;
            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (emitted >= maximumCells) yield break;
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    for (int row = firstRow; row <= lastRow; row++) {
                        for (int column = firstColumn; column <= lastColumn; column++) {
                            if (emitted >= maximumCells) yield break;
                            if (seen.Add(Key(row, column))) {
                                emitted++;
                                yield return (row, column);
                            }
                        }
                    }

                    continue;
                }

                if (A1.TryParseCellReferenceFast(token, out int singleRow, out int singleColumn) &&
                    seen.Add(Key(singleRow, singleColumn))) {
                    emitted++;
                    yield return (singleRow, singleColumn);
                }
            }
        }

        private static bool IsCellInReferenceList(int row, int column, string referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                        return true;
                    }
                } else if (A1.TryParseCellReferenceFast(token, out int singleRow, out int singleColumn)) {
                    (int Row, int Col) singleCell = (singleRow, singleColumn);
                    if (singleCell.Row == row && singleCell.Col == column) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string StripSheetPrefix(string reference) {
            int bang = reference.LastIndexOf('!');
            return bang >= 0 && bang + 1 < reference.Length ? reference.Substring(bang + 1) : reference;
        }

        private static string Key(int row, int column) => row.ToString(CultureInfo.InvariantCulture) + ":" + column.ToString(CultureInfo.InvariantCulture);

        private static bool TryGetConditionalNumericValue(string? text, out double numericValue) {
            if (!string.IsNullOrWhiteSpace(text) &&
                double.TryParse(text!.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue) &&
                !double.IsNaN(numericValue) &&
                !double.IsInfinity(numericValue)) {
                return true;
            }

            numericValue = 0D;
            return false;
        }

        private static bool TryNormalizeArgb(string? value, out string? argb) =>
            ExcelConditionalFormatThresholds.TryNormalizeArgb(value, out argb);

        private static bool HasSupportedDifferentialFormat(ExcelConditionalFormattingInfo rule) =>
            !string.IsNullOrWhiteSpace(rule.DifferentialFillColorArgb) ||
            !string.IsNullOrWhiteSpace(rule.DifferentialFontColorArgb) ||
            rule.DifferentialFontBold.HasValue ||
            rule.DifferentialFontItalic.HasValue ||
            rule.DifferentialFontUnderline.HasValue ||
            !string.IsNullOrWhiteSpace(rule.DifferentialFontName) ||
            rule.DifferentialFontSize.HasValue ||
            rule.DifferentialBorder != null;

        private static void ApplyDifferentialFormat(
            ExcelConditionalFormattingInfo rule,
            string key,
            Dictionary<string, ExcelConditionalCellFormat> formats) {
            if (!formats.TryGetValue(key, out ExcelConditionalCellFormat? format)) {
                format = new ExcelConditionalCellFormat();
                formats[key] = format;
            }

            if (string.IsNullOrWhiteSpace(format.FillColorArgb) && !string.IsNullOrWhiteSpace(rule.DifferentialFillColorArgb)) {
                format.FillColorArgb = rule.DifferentialFillColorArgb;
            }

            if (string.IsNullOrWhiteSpace(format.FontColorArgb) && !string.IsNullOrWhiteSpace(rule.DifferentialFontColorArgb)) {
                format.FontColorArgb = rule.DifferentialFontColorArgb;
            }

            format.FontBold ??= rule.DifferentialFontBold;
            format.FontItalic ??= rule.DifferentialFontItalic;
            format.FontUnderline ??= rule.DifferentialFontUnderline;

            if (string.IsNullOrWhiteSpace(format.FontName) && !string.IsNullOrWhiteSpace(rule.DifferentialFontName)) {
                format.FontName = rule.DifferentialFontName;
            }

            format.FontSize ??= rule.DifferentialFontSize;
            format.Border = MergeDifferentialBorder(format.Border, rule.DifferentialBorder);
        }

        private static ExcelCellBorderSnapshot? MergeDifferentialBorder(ExcelCellBorderSnapshot? current, ExcelCellBorderSnapshot? incoming) {
            if (incoming == null) {
                return current;
            }

            if (current == null) {
                return incoming;
            }

            return new ExcelCellBorderSnapshot {
                Left = current.Left ?? incoming.Left,
                Right = current.Right ?? incoming.Right,
                Top = current.Top ?? incoming.Top,
                Bottom = current.Bottom ?? incoming.Bottom,
                Diagonal = current.Diagonal ?? incoming.Diagonal,
                DiagonalUp = current.DiagonalUp || (current.Diagonal == null && incoming.DiagonalUp),
                DiagonalDown = current.DiagonalDown || (current.Diagonal == null && incoming.DiagonalDown)
            };
        }

        private static void ApplyFillFormat(
            string key,
            string fillColorArgb,
            Dictionary<string, ExcelConditionalCellFormat> formats) {
            if (!formats.TryGetValue(key, out ExcelConditionalCellFormat? format)) {
                format = new ExcelConditionalCellFormat();
                formats[key] = format;
            }

            format.FillColorArgb ??= fillColorArgb;
        }

        private readonly struct ConditionalNumericCell {
            internal ConditionalNumericCell(ExcelVisualCell cell, double value) {
                Cell = cell;
                Value = value;
            }

            internal ExcelVisualCell Cell { get; }

            internal double Value { get; }
        }
    }

    internal sealed class ExcelConditionalVisualState {
        internal static readonly ExcelConditionalVisualState Empty = new ExcelConditionalVisualState(
            new Dictionary<string, ExcelConditionalCellFormat>(StringComparer.Ordinal),
            Array.Empty<ExcelVisualConditionalDataBar>(),
            Array.Empty<ExcelVisualConditionalIcon>());

        internal ExcelConditionalVisualState(
            IReadOnlyDictionary<string, ExcelConditionalCellFormat> cellFormats,
            IReadOnlyList<ExcelVisualConditionalDataBar> dataBars,
            IReadOnlyList<ExcelVisualConditionalIcon> icons) {
            CellFormats = cellFormats;
            DataBars = dataBars;
            Icons = icons;
        }

        internal IReadOnlyDictionary<string, ExcelConditionalCellFormat> CellFormats { get; }

        internal IReadOnlyList<ExcelVisualConditionalDataBar> DataBars { get; }

        internal IReadOnlyList<ExcelVisualConditionalIcon> Icons { get; }
    }

    internal sealed class ExcelConditionalCellFormat {
        internal string? FillColorArgb { get; set; }

        internal string? FontColorArgb { get; set; }

        internal bool? FontBold { get; set; }

        internal bool? FontItalic { get; set; }

        internal bool? FontUnderline { get; set; }

        internal string? FontName { get; set; }

        internal double? FontSize { get; set; }

        internal ExcelCellBorderSnapshot? Border { get; set; }
    }
}
