using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private int _nextConditionalFormattingPriority;

        /// <summary>
        /// Lists conditional formatting rules on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelConditionalFormattingInfo> GetConditionalFormattingRules(string? a1Range = null) =>
            GetConditionalFormattingRules(a1Range, int.MaxValue, out _);

        internal IReadOnlyList<ExcelConditionalFormattingInfo> GetConditionalFormattingRules(
            string? a1Range,
            int maximumRules,
            out bool truncated) {
            if (maximumRules <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumRules));
            }

            truncated = false;
            (int r1, int c1, int r2, int c2)? filter = string.IsNullOrWhiteSpace(a1Range) ? null : ParseReferenceArgument(a1Range!);
            var workbookPart = _excelDocument.WorkbookPartRoot;
            Stylesheet? stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            var list = new List<ExcelConditionalFormattingInfo>();
            SortedSet<ConditionalFormattingCandidate>? retained = maximumRules == int.MaxValue
                ? null
                : new SortedSet<ConditionalFormattingCandidate>(ConditionalFormattingCandidateComparer.Instance);
            long ruleOrder = 0L;
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>()) {
                string range = conditional.SequenceOfReferences?.InnerText ?? string.Empty;
                if (filter.HasValue && !string.IsNullOrWhiteSpace(range)) {
                    if (!ReferenceListOverlaps(range, filter.Value)) continue;
                }

                foreach (var rule in conditional.Elements<ConditionalFormattingRule>()) {
                    if (retained == null) {
                        list.Add(ReadConditionalFormattingInfo(rule, range, stylesheet, workbookPart));
                        continue;
                    }

                    var candidate = new ConditionalFormattingCandidate(
                        rule,
                        range,
                        NormalizeConditionalFormattingPriority(rule),
                        ruleOrder++);
                    if (retained.Count < maximumRules) {
                        retained.Add(candidate);
                        continue;
                    }

                    truncated = true;
                    ConditionalFormattingCandidate worst = retained.Max;
                    if (ConditionalFormattingCandidateComparer.Instance.Compare(candidate, worst) < 0) {
                        retained.Remove(worst);
                        retained.Add(candidate);
                    }
                }
            }

            if (retained != null) {
                foreach (ConditionalFormattingCandidate candidate in retained) {
                    list.Add(ReadConditionalFormattingInfo(candidate.Rule, candidate.Range, stylesheet, workbookPart));
                }
            }

            return list;
        }

        private static ExcelConditionalFormattingInfo ReadConditionalFormattingInfo(
            ConditionalFormattingRule rule,
            string range,
            Stylesheet? stylesheet,
            WorkbookPart workbookPart) {
            uint? differentialFormatId = ReadDifferentialFormatId(rule);
            return new ExcelConditionalFormattingInfo {
                Range = range,
                Type = ReadConditionalFormatType(rule),
                Operator = ReadConditionalFormatOperator(rule),
                Text = rule.Text?.Value,
                TimePeriod = ReadConditionalTimePeriod(rule),
                Priority = (int)(rule.Priority?.Value ?? 0),
                StopIfTrue = rule.StopIfTrue?.Value ?? false,
                DifferentialFormatId = differentialFormatId,
                DifferentialFillColorArgb = ReadDifferentialFillColor(stylesheet, workbookPart, differentialFormatId),
                DifferentialFontColorArgb = ReadDifferentialFontColor(stylesheet, workbookPart, differentialFormatId),
                DifferentialFontBold = ReadDifferentialFontBold(stylesheet, differentialFormatId),
                DifferentialFontItalic = ReadDifferentialFontItalic(stylesheet, differentialFormatId),
                DifferentialFontUnderline = ReadDifferentialFontUnderline(stylesheet, differentialFormatId),
                DifferentialFontName = ReadDifferentialFontName(stylesheet, differentialFormatId),
                DifferentialFontSize = ReadDifferentialFontSize(stylesheet, differentialFormatId),
                DifferentialBorder = ReadDifferentialBorder(stylesheet, workbookPart, differentialFormatId),
                Formulas = rule.Elements<Formula>().Select(f => f.Text ?? string.Empty).ToArray(),
                ColorScaleColors = ReadColorScaleColors(rule),
                ColorScaleThresholds = ReadColorScaleThresholds(rule),
                DataBarColor = ReadDataBarColor(rule),
                DataBarThresholds = ReadDataBarThresholds(rule),
                DataBarShowValue = ReadDataBarShowValue(rule),
                IconSet = ReadIconSetName(rule),
                IconSetShowValue = ReadIconSetShowValue(rule),
                IconSetReverse = ReadIconSetReverse(rule),
                IconSetThresholds = ReadIconSetThresholds(rule),
                TopBottomRank = rule.Rank?.Value,
                TopBottomBottom = rule.Bottom?.Value ?? false,
                TopBottomPercent = rule.Percent?.Value ?? false,
                AboveAverageAbove = rule.AboveAverage?.Value ?? true,
                AboveAverageEqual = rule.EqualAverage?.Value ?? false,
                AboveAverageStdDev = rule.StdDev?.Value
            };
        }

        private static int NormalizeConditionalFormattingPriority(ConditionalFormattingRule rule) {
            int priority = (int)(rule.Priority?.Value ?? 0);
            return priority <= 0 ? int.MaxValue : priority;
        }

        private readonly struct ConditionalFormattingCandidate {
            internal ConditionalFormattingCandidate(
                ConditionalFormattingRule rule,
                string range,
                int priority,
                long order) {
                Rule = rule;
                Range = range;
                Priority = priority;
                Order = order;
            }

            internal ConditionalFormattingRule Rule { get; }
            internal string Range { get; }
            internal int Priority { get; }
            internal long Order { get; }
        }

        private sealed class ConditionalFormattingCandidateComparer : IComparer<ConditionalFormattingCandidate> {
            internal static readonly ConditionalFormattingCandidateComparer Instance = new ConditionalFormattingCandidateComparer();

            public int Compare(ConditionalFormattingCandidate left, ConditionalFormattingCandidate right) {
                int priority = left.Priority.CompareTo(right.Priority);
                return priority != 0 ? priority : left.Order.CompareTo(right.Order);
            }
        }

        private static string ReadConditionalFormatType(ConditionalFormattingRule rule) {
            if (rule.Type == null) {
                return string.Empty;
            }

            ConditionalFormatValues value = rule.Type.Value;
            if (value == ConditionalFormatValues.CellIs) return nameof(ConditionalFormatValues.CellIs);
            if (value == ConditionalFormatValues.Expression) return nameof(ConditionalFormatValues.Expression);
            if (value == ConditionalFormatValues.ColorScale) return nameof(ConditionalFormatValues.ColorScale);
            if (value == ConditionalFormatValues.DataBar) return nameof(ConditionalFormatValues.DataBar);
            if (value == ConditionalFormatValues.IconSet) return nameof(ConditionalFormatValues.IconSet);
            if (value == ConditionalFormatValues.Top10) return nameof(ConditionalFormatValues.Top10);
            if (value == ConditionalFormatValues.UniqueValues) return nameof(ConditionalFormatValues.UniqueValues);
            if (value == ConditionalFormatValues.DuplicateValues) return nameof(ConditionalFormatValues.DuplicateValues);
            if (value == ConditionalFormatValues.ContainsText) return nameof(ConditionalFormatValues.ContainsText);
            if (value == ConditionalFormatValues.NotContainsText) return nameof(ConditionalFormatValues.NotContainsText);
            if (value == ConditionalFormatValues.BeginsWith) return nameof(ConditionalFormatValues.BeginsWith);
            if (value == ConditionalFormatValues.EndsWith) return nameof(ConditionalFormatValues.EndsWith);
            if (value == ConditionalFormatValues.ContainsBlanks) return nameof(ConditionalFormatValues.ContainsBlanks);
            if (value == ConditionalFormatValues.NotContainsBlanks) return nameof(ConditionalFormatValues.NotContainsBlanks);
            if (value == ConditionalFormatValues.ContainsErrors) return nameof(ConditionalFormatValues.ContainsErrors);
            if (value == ConditionalFormatValues.NotContainsErrors) return nameof(ConditionalFormatValues.NotContainsErrors);
            if (value == ConditionalFormatValues.TimePeriod) return nameof(ConditionalFormatValues.TimePeriod);
            if (value == ConditionalFormatValues.AboveAverage) return nameof(ConditionalFormatValues.AboveAverage);

            return rule.Type.InnerText ?? string.Empty;
        }

        private static string? ReadConditionalFormatOperator(ConditionalFormattingRule rule) {
            if (rule.Operator == null) {
                return null;
            }

            ConditionalFormattingOperatorValues value = rule.Operator.Value;
            if (value == ConditionalFormattingOperatorValues.Between) return nameof(ConditionalFormattingOperatorValues.Between);
            if (value == ConditionalFormattingOperatorValues.NotBetween) return nameof(ConditionalFormattingOperatorValues.NotBetween);
            if (value == ConditionalFormattingOperatorValues.Equal) return nameof(ConditionalFormattingOperatorValues.Equal);
            if (value == ConditionalFormattingOperatorValues.NotEqual) return nameof(ConditionalFormattingOperatorValues.NotEqual);
            if (value == ConditionalFormattingOperatorValues.GreaterThan) return nameof(ConditionalFormattingOperatorValues.GreaterThan);
            if (value == ConditionalFormattingOperatorValues.LessThan) return nameof(ConditionalFormattingOperatorValues.LessThan);
            if (value == ConditionalFormattingOperatorValues.GreaterThanOrEqual) return nameof(ConditionalFormattingOperatorValues.GreaterThanOrEqual);
            if (value == ConditionalFormattingOperatorValues.LessThanOrEqual) return nameof(ConditionalFormattingOperatorValues.LessThanOrEqual);
            if (value == ConditionalFormattingOperatorValues.ContainsText) return nameof(ConditionalFormattingOperatorValues.ContainsText);
            if (value == ConditionalFormattingOperatorValues.NotContains) return nameof(ConditionalFormattingOperatorValues.NotContains);
            if (value == ConditionalFormattingOperatorValues.BeginsWith) return nameof(ConditionalFormattingOperatorValues.BeginsWith);
            if (value == ConditionalFormattingOperatorValues.EndsWith) return nameof(ConditionalFormattingOperatorValues.EndsWith);

            return rule.Operator.InnerText;
        }

        private static string? ReadConditionalTimePeriod(ConditionalFormattingRule rule) {
            if (rule.TimePeriod == null) {
                return null;
            }

            TimePeriodValues value = rule.TimePeriod.Value;
            if (value == TimePeriodValues.Yesterday) return nameof(TimePeriodValues.Yesterday);
            if (value == TimePeriodValues.Today) return nameof(TimePeriodValues.Today);
            if (value == TimePeriodValues.Tomorrow) return nameof(TimePeriodValues.Tomorrow);
            if (value == TimePeriodValues.Last7Days) return nameof(TimePeriodValues.Last7Days);
            if (value == TimePeriodValues.LastWeek) return nameof(TimePeriodValues.LastWeek);
            if (value == TimePeriodValues.ThisWeek) return nameof(TimePeriodValues.ThisWeek);
            if (value == TimePeriodValues.NextWeek) return nameof(TimePeriodValues.NextWeek);
            if (value == TimePeriodValues.LastMonth) return nameof(TimePeriodValues.LastMonth);
            if (value == TimePeriodValues.ThisMonth) return nameof(TimePeriodValues.ThisMonth);
            if (value == TimePeriodValues.NextMonth) return nameof(TimePeriodValues.NextMonth);

            return rule.TimePeriod.InnerText;
        }

        private static IReadOnlyList<string> ReadColorScaleColors(ConditionalFormattingRule rule) {
            ColorScale? colorScale = rule.GetFirstChild<ColorScale>();
            if (colorScale == null) {
                return Array.Empty<string>();
            }

            return colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>()
                .Select(color => color.Rgb?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToArray();
        }

        private static IReadOnlyList<ExcelConditionalFormatThreshold> ReadColorScaleThresholds(ConditionalFormattingRule rule) {
            ColorScale? colorScale = rule.GetFirstChild<ColorScale>();
            return ReadConditionalFormatThresholds(colorScale);
        }

        private static string? ReadDataBarColor(ConditionalFormattingRule rule) {
            DataBar? dataBar = rule.GetFirstChild<DataBar>();
            if (dataBar == null) {
                return null;
            }

            return dataBar.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>()
                .Select(color => color.Rgb?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        private static bool ReadDataBarShowValue(ConditionalFormattingRule rule) {
            DataBar? dataBar = rule.GetFirstChild<DataBar>();
            return dataBar?.ShowValue?.Value ?? true;
        }

        private static IReadOnlyList<ExcelConditionalFormatThreshold> ReadDataBarThresholds(ConditionalFormattingRule rule) {
            DataBar? dataBar = rule.GetFirstChild<DataBar>();
            return ReadConditionalFormatThresholds(dataBar);
        }

        private static IReadOnlyList<ExcelConditionalFormatThreshold> ReadConditionalFormatThresholds(OpenXmlElement? parent) {
            if (parent == null) {
                return Array.Empty<ExcelConditionalFormatThreshold>();
            }

            return parent.Elements<ConditionalFormatValueObject>()
                .Select(threshold => new ExcelConditionalFormatThreshold {
                    Type = threshold.Type?.InnerText ?? string.Empty,
                    Value = threshold.Val?.Value
                })
                .ToArray();
        }

        private static string? ReadIconSetName(ConditionalFormattingRule rule) {
            IconSet? iconSet = rule.GetFirstChild<IconSet>();
            if (iconSet?.IconSetValue?.Value == IconSetValues.ThreeTrafficLights1) {
                return nameof(IconSetValues.ThreeTrafficLights1);
            }

            return iconSet?.IconSetValue?.InnerText;
        }

        private static bool ReadIconSetShowValue(ConditionalFormattingRule rule) {
            IconSet? iconSet = rule.GetFirstChild<IconSet>();
            return iconSet?.ShowValue?.Value ?? true;
        }

        private static bool ReadIconSetReverse(ConditionalFormattingRule rule) {
            IconSet? iconSet = rule.GetFirstChild<IconSet>();
            return iconSet?.Reverse?.Value ?? false;
        }

        private static IReadOnlyList<ExcelConditionalIconSetThreshold> ReadIconSetThresholds(ConditionalFormattingRule rule) {
            IconSet? iconSet = rule.GetFirstChild<IconSet>();
            if (iconSet == null) {
                return Array.Empty<ExcelConditionalIconSetThreshold>();
            }

            return iconSet.Elements<ConditionalFormatValueObject>()
                .Select(threshold => new ExcelConditionalIconSetThreshold {
                    Type = threshold.Type?.InnerText ?? string.Empty,
                    Value = threshold.Val?.Value,
                    GreaterThanOrEqual = threshold.GreaterThanOrEqual?.Value ?? true
                })
                .ToArray();
        }

        /// <summary>
        /// Clears conditional formatting rules, optionally restricted to a range.
        /// </summary>
        public void ClearConditionalFormatting(string? a1Range = null) {
            WriteLock(() => ClearConditionalFormattingCore(a1Range));
        }

        /// <summary>
        /// Adds a formula-based conditional formatting rule.
        /// </summary>
        public void AddConditionalFormulaRule(string range, string formula) {
            AddConditionalFormulaRule(range, formula, stopIfTrue: false, fillColor: null, priority: null);
        }

        /// <summary>
        /// Adds a formula-based conditional formatting rule.
        /// </summary>
        public void AddConditionalFormulaRule(string range, string formula, bool stopIfTrue) {
            AddConditionalFormulaRule(range, formula, stopIfTrue, fillColor: null, priority: null);
        }

        /// <summary>
        /// Adds a formula-based conditional formatting rule.
        /// </summary>
        public void AddConditionalFormulaRule(string range, string formula, bool stopIfTrue, int? priority) {
            AddConditionalFormulaRule(range, formula, stopIfTrue, fillColor: null, priority);
        }

        /// <summary>
        /// Adds a formula-based conditional formatting rule.
        /// </summary>
        public void AddConditionalFormulaRule(string range, string formula, bool stopIfTrue = false, string? fillColor = null, int? priority = null) {
            AddConditionalRuleCore(range, ConditionalFormatValues.Expression, null, new[] { formula }, stopIfTrue, fillColor, priority);
        }

        /// <summary>
        /// Adds a duplicate-values conditional formatting rule.
        /// </summary>
        public void AddConditionalDuplicateValuesRule(string range) {
            AddConditionalDuplicateValuesRule(range, fillColor: null);
        }

        /// <summary>
        /// Adds a duplicate-values conditional formatting rule with an optional solid fill differential format.
        /// </summary>
        public void AddConditionalDuplicateValuesRule(string range, string? fillColor) {
            AddConditionalRuleCore(range, ConditionalFormatValues.DuplicateValues, null, Array.Empty<string>(), stopIfTrue: false, fillColor: fillColor, priority: null);
        }

        /// <summary>
        /// Adds a unique-values conditional formatting rule.
        /// </summary>
        public void AddConditionalUniqueValuesRule(string range) {
            AddConditionalUniqueValuesRule(range, fillColor: null);
        }

        /// <summary>
        /// Adds a unique-values conditional formatting rule with an optional solid fill differential format.
        /// </summary>
        public void AddConditionalUniqueValuesRule(string range, string? fillColor) {
            AddConditionalRuleCore(range, ConditionalFormatValues.UniqueValues, null, Array.Empty<string>(), stopIfTrue: false, fillColor: fillColor, priority: null);
        }

        /// <summary>
        /// Adds an above/below-average conditional formatting rule with optional standard-deviation, stop-if-true, and solid fill settings.
        /// </summary>
        public void AddConditionalAboveAverageRule(
            string range,
            bool aboveAverage = true,
            bool equalAverage = false,
            uint? standardDeviation = null,
            bool stopIfTrue = false,
            string? fillColor = null) {
            AddConditionalRuleCore(range, ConditionalFormatValues.AboveAverage, rule => {
                rule.AboveAverage = aboveAverage;
                rule.EqualAverage = equalAverage;
                if (standardDeviation.HasValue) {
                    rule.StdDev = checked((int)standardDeviation.Value);
                }
            }, Array.Empty<string>(), stopIfTrue, fillColor: fillColor);
        }

        /// <summary>
        /// Adds a contains-text conditional formatting rule.
        /// </summary>
        public void AddConditionalContainsTextRule(string range, string text, string? fillColor = null) {
            AddConditionalTextRule(range, text, ConditionalFormatValues.ContainsText, ConditionalFormattingOperatorValues.ContainsText, fillColor);
        }

        /// <summary>
        /// Adds a not-contains-text conditional formatting rule.
        /// </summary>
        public void AddConditionalNotContainsTextRule(string range, string text, string? fillColor = null) {
            AddConditionalTextRule(range, text, ConditionalFormatValues.NotContainsText, ConditionalFormattingOperatorValues.NotContains, fillColor);
        }

        /// <summary>
        /// Adds a begins-with conditional formatting rule.
        /// </summary>
        public void AddConditionalBeginsWithTextRule(string range, string text, string? fillColor = null) {
            AddConditionalTextRule(range, text, ConditionalFormatValues.BeginsWith, ConditionalFormattingOperatorValues.BeginsWith, fillColor);
        }

        /// <summary>
        /// Adds an ends-with conditional formatting rule.
        /// </summary>
        public void AddConditionalEndsWithTextRule(string range, string text, string? fillColor = null) {
            AddConditionalTextRule(range, text, ConditionalFormatValues.EndsWith, ConditionalFormattingOperatorValues.EndsWith, fillColor);
        }

        /// <summary>
        /// Adds a top/bottom conditional formatting rule.
        /// </summary>
        public void AddConditionalTopBottomRule(string range, uint rank, bool bottom = false, bool percent = false) {
            AddConditionalTopBottomRule(range, rank, bottom, percent, fillColor: null);
        }

        /// <summary>
        /// Adds a top/bottom conditional formatting rule with an optional solid fill differential format.
        /// </summary>
        public void AddConditionalTopBottomRule(string range, uint rank, bool bottom, bool percent, string? fillColor) {
            if (rank == 0) throw new ArgumentOutOfRangeException(nameof(rank));
            AddConditionalRuleCore(range, ConditionalFormatValues.Top10, rule => {
                rule.Rank = rank;
                rule.Bottom = bottom;
                rule.Percent = percent;
            }, Array.Empty<string>(), stopIfTrue: false, fillColor: fillColor);
        }

        private void AddConditionalTextRule(
            string range,
            string text,
            ConditionalFormatValues type,
            ConditionalFormattingOperatorValues @operator,
            string? fillColor) {
            if (string.IsNullOrWhiteSpace(text)) {
                throw new ArgumentException("Conditional text cannot be empty.", nameof(text));
            }

            string firstCell = GetFirstReferenceCell(range);
            AddConditionalRuleCore(range, type, rule => {
                rule.Operator = @operator;
                rule.Text = text;
            }, new[] { BuildTextRuleFormula(firstCell, text, type) }, stopIfTrue: false, fillColor: fillColor);
        }

        private static string BuildTextRuleFormula(string cellReference, string text, ConditionalFormatValues type) {
            string escaped = EscapeExcelString(text);
            if (type == ConditionalFormatValues.ContainsText) {
                return "NOT(ISERROR(SEARCH(\"" + escaped + "\"," + cellReference + ")))";
            }

            if (type == ConditionalFormatValues.NotContainsText) {
                return "ISERROR(SEARCH(\"" + escaped + "\"," + cellReference + "))";
            }

            if (type == ConditionalFormatValues.BeginsWith) {
                return "LEFT(" + cellReference + ",LEN(\"" + escaped + "\"))=\"" + escaped + "\"";
            }

            return "RIGHT(" + cellReference + ",LEN(\"" + escaped + "\"))=\"" + escaped + "\"";
        }

        private static string EscapeExcelString(string text) => text.Replace("\"", "\"\"");

        private static string GetFirstReferenceCell(string range) {
            var bounds = ParseReferenceArgument(range);
            return A1.ColumnIndexToLetters(bounds.c1) + bounds.r1.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Adds a text-matching conditional formatting rule.
        /// </summary>
        public void AddConditionalTextRule(string range, ConditionalFormatValues type, string text, bool stopIfTrue = false) {
            if (string.IsNullOrEmpty(text)) throw new ArgumentNullException(nameof(text));
            string firstCell = GetFirstCellReference(range);
            string literal = EscapeFormulaString(text);
            string formula;
            ConditionalFormattingOperatorValues op;
            if (type == ConditionalFormatValues.ContainsText) {
                formula = $"NOT(ISERROR(SEARCH(\"{literal}\",{firstCell})))";
                op = ConditionalFormattingOperatorValues.ContainsText;
            } else if (type == ConditionalFormatValues.NotContainsText) {
                formula = $"ISERROR(SEARCH(\"{literal}\",{firstCell}))";
                op = ConditionalFormattingOperatorValues.NotContains;
            } else if (type == ConditionalFormatValues.BeginsWith) {
                formula = $"LEFT({firstCell},LEN(\"{literal}\"))=\"{literal}\"";
                op = ConditionalFormattingOperatorValues.BeginsWith;
            } else if (type == ConditionalFormatValues.EndsWith) {
                formula = $"RIGHT({firstCell},LEN(\"{literal}\"))=\"{literal}\"";
                op = ConditionalFormattingOperatorValues.EndsWith;
            } else {
                throw new ArgumentOutOfRangeException(nameof(type), "Text conditional formatting requires ContainsText, NotContainsText, BeginsWith, or EndsWith.");
            }

            AddConditionalRuleCore(range, type, rule => {
                rule.Text = text;
                rule.Operator = op;
            }, new[] { formula }, stopIfTrue, fillColor: null);
        }

        /// <summary>
        /// Adds a blanks/non-blanks conditional formatting rule.
        /// </summary>
        public void AddConditionalBlanksRule(string range, bool containsBlanks = true, bool stopIfTrue = false) {
            string firstCell = GetFirstCellReference(range);
            ConditionalFormatValues type = containsBlanks ? ConditionalFormatValues.ContainsBlanks : ConditionalFormatValues.NotContainsBlanks;
            string formula = containsBlanks ? $"LEN(TRIM({firstCell}))=0" : $"LEN(TRIM({firstCell}))>0";
            AddConditionalRuleCore(range, type, null, new[] { formula }, stopIfTrue, fillColor: null);
        }

        /// <summary>
        /// Adds an errors/non-errors conditional formatting rule.
        /// </summary>
        public void AddConditionalErrorsRule(string range, bool containsErrors = true, bool stopIfTrue = false) {
            string firstCell = GetFirstCellReference(range);
            ConditionalFormatValues type = containsErrors ? ConditionalFormatValues.ContainsErrors : ConditionalFormatValues.NotContainsErrors;
            string formula = containsErrors ? $"ISERROR({firstCell})" : $"NOT(ISERROR({firstCell}))";
            AddConditionalRuleCore(range, type, null, new[] { formula }, stopIfTrue, fillColor: null);
        }

        /// <summary>
        /// Adds a time-period conditional formatting rule.
        /// </summary>
        public void AddConditionalTimePeriodRule(string range, TimePeriodValues timePeriod, bool stopIfTrue = false, string? fillColor = null) {
            string firstCell = GetFirstCellReference(range);
            AddConditionalRuleCore(range, ConditionalFormatValues.TimePeriod, rule => {
                rule.TimePeriod = timePeriod;
            }, new[] { BuildTimePeriodFormula(firstCell, timePeriod) }, stopIfTrue, fillColor);
        }

        /// <summary>
        /// Lists data validation rules on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelDataValidationInfo> GetDataValidations(string? a1Range = null) {
            (int r1, int c1, int r2, int c2)? filter = string.IsNullOrWhiteSpace(a1Range) ? null : ParseReferenceArgument(a1Range!);
            var result = new List<ExcelDataValidationInfo>();
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) return result;

            foreach (var validation in validations.Elements<DataValidation>()) {
                string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                if (filter.HasValue) {
                    if (!ReferenceListOverlaps(range, filter.Value)) continue;
                }

                result.Add(new ExcelDataValidationInfo {
                    Range = range,
                    Type = validation.Type?.InnerText ?? string.Empty,
                    Operator = validation.Operator?.InnerText,
                    AllowBlank = validation.AllowBlank?.Value ?? false,
                    SuppressDropDown = validation.ShowDropDown?.Value ?? false,
                    ErrorStyle = validation.ErrorStyle?.InnerText,
                    ShowInputMessage = validation.ShowInputMessage?.Value ?? false,
                    ShowErrorMessage = validation.ShowErrorMessage?.Value ?? false,
                    Formula1 = validation.GetFirstChild<Formula1>()?.Text,
                    Formula2 = validation.GetFirstChild<Formula2>()?.Text,
                    PromptTitle = validation.PromptTitle?.Value,
                    Prompt = validation.Prompt?.Value,
                    ErrorTitle = validation.ErrorTitle?.Value,
                    Error = validation.Error?.Value
                });
            }

            return result;
        }

        /// <summary>
        /// Removes data validation rules, optionally restricted to a range.
        /// </summary>
        public void RemoveDataValidations(string? a1Range = null) {
            WriteLock(() => RemoveDataValidationsCore(a1Range));
        }

        /// <summary>
        /// Applies prompt/error-message metadata to existing validations that overlap the range.
        /// </summary>
        public void SetDataValidationMessages(string a1Range, ExcelDataValidationMessageOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            var filter = ParseReferenceArgument(a1Range);
            WriteLock(() => {
                var validations = WorksheetRoot.GetFirstChild<DataValidations>();
                if (validations == null) return;
                bool changed = false;
                bool showInputMessage = options.PreserveShowMessageFlags
                    ? options.ShowInputMessage
                    : options.ShowInputMessage || !string.IsNullOrEmpty(options.Prompt) || !string.IsNullOrEmpty(options.PromptTitle);
                bool showErrorMessage = options.PreserveShowMessageFlags
                    ? options.ShowErrorMessage
                    : options.ShowErrorMessage || !string.IsNullOrEmpty(options.Error) || !string.IsNullOrEmpty(options.ErrorTitle);
                foreach (var validation in validations.Elements<DataValidation>()) {
                    string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                    if (!ReferenceListOverlaps(range, filter)) continue;

                    bool validationChanged =
                        !string.Equals(validation.PromptTitle?.Value, options.PromptTitle, StringComparison.Ordinal)
                        || !string.Equals(validation.Prompt?.Value, options.Prompt, StringComparison.Ordinal)
                        || !string.Equals(validation.ErrorTitle?.Value, options.ErrorTitle, StringComparison.Ordinal)
                        || !string.Equals(validation.Error?.Value, options.Error, StringComparison.Ordinal)
                        || validation.ShowInputMessage?.Value != showInputMessage
                        || validation.ShowErrorMessage?.Value != showErrorMessage
                        || (options.ErrorStyle.HasValue && validation.ErrorStyle?.Value != options.ErrorStyle.Value)
                        || (options.SuppressDropDown.HasValue && validation.ShowDropDown?.Value != options.SuppressDropDown.Value);

                    if (!validationChanged) continue;

                    validation.PromptTitle = options.PromptTitle;
                    validation.Prompt = options.Prompt;
                    validation.ErrorTitle = options.ErrorTitle;
                    validation.Error = options.Error;
                    validation.ShowInputMessage = showInputMessage;
                    validation.ShowErrorMessage = showErrorMessage;
                    if (options.ErrorStyle.HasValue) {
                        validation.ErrorStyle = options.ErrorStyle.Value;
                    }

                    if (options.SuppressDropDown.HasValue) {
                        validation.ShowDropDown = options.SuppressDropDown.Value;
                    }

                    changed = true;
                }

                if (changed) {
                    WorksheetRoot.Save();
                }
            });
        }

        private void AddConditionalRuleCore(string range, ConditionalFormatValues type, Action<ConditionalFormattingRule>? configure, IReadOnlyList<string> formulas, bool stopIfTrue, string? fillColor = null, int? priority = null) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                _excelDocument.EnsureWorkbookThemeAndStyles();
                Worksheet worksheet = WorksheetRoot;
                var conditional = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };
                var rule = new ConditionalFormattingRule {
                    Type = type,
                    Priority = priority ?? GetNextConditionalFormattingPriority(),
                    StopIfTrue = stopIfTrue
                };
                configure?.Invoke(rule);
                if (!string.IsNullOrWhiteSpace(fillColor)) {
                    rule.FormatId = GetOrCreateDifferentialFillFormatId(fillColor!);
                }

                foreach (var formula in formulas) {
                    rule.Append(new Formula(formula));
                }
                conditional.Append(rule);
                InsertConditionalFormatting(conditional);
            });
        }

        private uint GetOrCreateDifferentialFillFormatId(string fillColor) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesPart.Stylesheet);

            string argb = NormalizeHexColor(fillColor);
            var candidate = new DifferentialFormat(
                new Fill(
                    new PatternFill {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = argb },
                        BackgroundColor = new BackgroundColor { Rgb = argb }
                    }));

            DifferentialFormats differentialFormats = stylesPart.Stylesheet.DifferentialFormats ??= new DifferentialFormats();
            var existing = differentialFormats.Elements<DifferentialFormat>()
                .Select((format, index) => new { format, index })
                .FirstOrDefault(entry => string.Equals(entry.format.OuterXml, candidate.OuterXml, StringComparison.Ordinal));
            if (existing != null) {
                return (uint)existing.index;
            }

            differentialFormats.Append(candidate);
            differentialFormats.Count = (uint)differentialFormats.Count();
            stylesPart.Stylesheet.Save();
            return differentialFormats.Count!.Value - 1U;
        }

        private static uint? ReadDifferentialFormatId(ConditionalFormattingRule rule) {
            return rule.FormatId?.Value;
        }

        private static string? ReadDifferentialFillColor(Stylesheet? stylesheet, DocumentFormat.OpenXml.Packaging.WorkbookPart? workbookPart, uint? differentialFormatId) {
            if (!differentialFormatId.HasValue || stylesheet?.DifferentialFormats == null) {
                return null;
            }

            DifferentialFormat? format = stylesheet.DifferentialFormats.Elements<DifferentialFormat>().ElementAtOrDefault((int)differentialFormatId.Value);
            PatternFill? pattern = format?.Fill?.PatternFill;
            if (pattern == null || pattern.PatternType?.Value != PatternValues.Solid) {
                return null;
            }

            return ExcelThemeColorResolver.Resolve(pattern.ForegroundColor, workbookPart)
                ?? ExcelThemeColorResolver.Resolve(pattern.BackgroundColor, workbookPart);
        }

        private static string? ReadDifferentialFontColor(Stylesheet? stylesheet, DocumentFormat.OpenXml.Packaging.WorkbookPart? workbookPart, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            return ExcelThemeColorResolver.Resolve(format?.Font?.Color, workbookPart);
        }

        private static bool? ReadDifferentialFontBold(Stylesheet? stylesheet, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            return ReadDifferentialBoolean(format?.Font?.Bold);
        }

        private static bool? ReadDifferentialFontItalic(Stylesheet? stylesheet, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            return ReadDifferentialBoolean(format?.Font?.Italic);
        }

        private static bool? ReadDifferentialFontUnderline(Stylesheet? stylesheet, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            return ReadDifferentialBoolean(format?.Font?.Underline);
        }

        private static bool? ReadDifferentialBoolean(OpenXmlElement? element) {
            if (element == null) {
                return null;
            }

            string? value = element.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "val", StringComparison.OrdinalIgnoreCase))
                .Value;
            if (string.IsNullOrWhiteSpace(value)) {
                return true;
            }

            if (string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return true;
        }

        private static string? ReadDifferentialFontName(Stylesheet? stylesheet, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            string? fontName = format?.Font?.FontName?.Val?.Value;
            return string.IsNullOrWhiteSpace(fontName) ? null : fontName;
        }

        private static double? ReadDifferentialFontSize(Stylesheet? stylesheet, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            double? fontSize = format?.Font?.FontSize?.Val?.Value;
            return fontSize is > 0D ? fontSize : null;
        }

        private static ExcelCellBorderSnapshot? ReadDifferentialBorder(Stylesheet? stylesheet, WorkbookPart? workbookPart, uint? differentialFormatId) {
            DifferentialFormat? format = GetDifferentialFormat(stylesheet, differentialFormatId);
            ExcelCellBorderSnapshot? border = BuildBorderSnapshot(format?.Border, workbookPart);
            return RemoveDifferentialColorOnlyBorderSides(border);
        }

        private static ExcelCellBorderSnapshot? RemoveDifferentialColorOnlyBorderSides(ExcelCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            ExcelBorderSideSnapshot? left = HasBorderStyle(border.Left) ? border.Left : null;
            ExcelBorderSideSnapshot? right = HasBorderStyle(border.Right) ? border.Right : null;
            ExcelBorderSideSnapshot? top = HasBorderStyle(border.Top) ? border.Top : null;
            ExcelBorderSideSnapshot? bottom = HasBorderStyle(border.Bottom) ? border.Bottom : null;
            ExcelBorderSideSnapshot? diagonal = HasBorderStyle(border.Diagonal) ? border.Diagonal : null;
            bool diagonalUp = diagonal != null && border.DiagonalUp;
            bool diagonalDown = diagonal != null && border.DiagonalDown;
            if (left == null && right == null && top == null && bottom == null && diagonal == null) {
                return null;
            }

            return new ExcelCellBorderSnapshot {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Diagonal = diagonal,
                DiagonalUp = diagonalUp,
                DiagonalDown = diagonalDown
            };
        }

        private static bool HasBorderStyle(ExcelBorderSideSnapshot? side) {
            if (side == null || string.IsNullOrWhiteSpace(side.Style)) {
                return false;
            }

            return !string.Equals(side.Style, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static DifferentialFormat? GetDifferentialFormat(Stylesheet? stylesheet, uint? differentialFormatId) {
            if (!differentialFormatId.HasValue || stylesheet?.DifferentialFormats == null) {
                return null;
            }

            return stylesheet.DifferentialFormats.Elements<DifferentialFormat>().ElementAtOrDefault((int)differentialFormatId.Value);
        }

        internal uint AppendConditionalDifferentialFormat(DifferentialFormat differentialFormat) {
            if (differentialFormat == null) throw new ArgumentNullException(nameof(differentialFormat));
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            uint id = (uint)stylesheet.DifferentialFormats.Elements<DifferentialFormat>().Count();
            stylesheet.DifferentialFormats.Append(differentialFormat);
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Elements<DifferentialFormat>().Count();
            stylesPart.Stylesheet.Save();
            return id;
        }

        internal void SetLastConditionalFormattingRuleDifferentialFormatId(uint differentialFormatId) {
            ConditionalFormattingRule? rule = WorksheetRoot.Elements<ConditionalFormatting>()
                .LastOrDefault()?
                .Elements<ConditionalFormattingRule>()
                .LastOrDefault();
            if (rule == null) {
                return;
            }

            rule.FormatId = differentialFormatId;
        }

        internal void SetConditionalFormattingRuleDifferentialFormatId(string range, int? priority, ConditionalFormatValues type, uint differentialFormatId) {
            ConditionalFormattingRule? rule = WorksheetRoot.Elements<ConditionalFormatting>()
                .Where(formatting => string.Equals(formatting.SequenceOfReferences?.InnerText, range, StringComparison.Ordinal))
                .SelectMany(formatting => formatting.Elements<ConditionalFormattingRule>())
                .Where(rule => rule.Type?.Value == type)
                .Where(rule => !priority.HasValue || rule.Priority?.Value == priority.Value)
                .LastOrDefault();
            if (rule == null) {
                return;
            }

            rule.FormatId = differentialFormatId;
        }

        private static string GetFirstCellReference(string range) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            string firstReference = range.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries)[0];
            int sheetSeparatorIndex = firstReference.LastIndexOf('!');
            if (sheetSeparatorIndex >= 0 && sheetSeparatorIndex < firstReference.Length - 1) {
                firstReference = firstReference.Substring(sheetSeparatorIndex + 1);
            }

            int rangeSeparatorIndex = firstReference.IndexOf(':');
            if (rangeSeparatorIndex >= 0) {
                firstReference = firstReference.Substring(0, rangeSeparatorIndex);
            }

            return firstReference.Replace("$", string.Empty);
        }

        private static string EscapeFormulaString(string value) {
            return value.Replace("\"", "\"\"");
        }

        private static string BuildTimePeriodFormula(string firstCell, TimePeriodValues timePeriod) {
            string day = $"FLOOR({firstCell},1)";
            if (timePeriod == TimePeriodValues.Yesterday) return $"{day}=TODAY()-1";
            if (timePeriod == TimePeriodValues.Today) return $"{day}=TODAY()";
            if (timePeriod == TimePeriodValues.Tomorrow) return $"{day}=TODAY()+1";
            if (timePeriod == TimePeriodValues.Last7Days) return $"AND(TODAY()-FLOOR({firstCell},1)<=6,FLOOR({firstCell},1)<=TODAY())";
            if (timePeriod == TimePeriodValues.LastWeek) return $"AND({day}>=TODAY()-WEEKDAY(TODAY(),2)-6,{day}<=TODAY()-WEEKDAY(TODAY(),2))";
            if (timePeriod == TimePeriodValues.ThisWeek) return $"AND({day}>=TODAY()-WEEKDAY(TODAY(),2)+1,{day}<=TODAY()-WEEKDAY(TODAY(),2)+7)";
            if (timePeriod == TimePeriodValues.NextWeek) return $"AND({day}>=TODAY()-WEEKDAY(TODAY(),2)+8,{day}<=TODAY()-WEEKDAY(TODAY(),2)+14)";
            if (timePeriod == TimePeriodValues.LastMonth) return $"AND({day}>=DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),{day}<DATE(YEAR(TODAY()),MONTH(TODAY()),1))";
            if (timePeriod == TimePeriodValues.ThisMonth) return $"AND({day}>=DATE(YEAR(TODAY()),MONTH(TODAY()),1),{day}<DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))";
            if (timePeriod == TimePeriodValues.NextMonth) return $"AND({day}>=DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),{day}<DATE(YEAR(TODAY()),MONTH(TODAY())+2,1))";

            return $"{day}=TODAY()";
        }

        private void ClearConditionalFormattingCore(string? a1Range) {
            bool changed = false;
            if (string.IsNullOrWhiteSpace(a1Range)) {
                foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                    conditional.Remove();
                    changed = true;
                }
            } else {
                var filter = ParseReferenceArgument(a1Range!);
                foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                    string range = conditional.SequenceOfReferences?.InnerText ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(range, filter, out var remaining)) {
                        continue;
                    }

                    if (remaining.Count == 0) {
                        conditional.Remove();
                        changed = true;
                    } else {
                        string replacement = string.Join(" ", remaining);
                        conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = replacement };
                        changed = true;
                    }
                }
            }

            if (changed) {
                _nextConditionalFormattingPriority = 0;
                WorksheetRoot.Save();
            }
        }

        private void RemoveDataValidationsCore(string? a1Range) {
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) return;

            bool changed = false;
            if (string.IsNullOrWhiteSpace(a1Range)) {
                foreach (var validation in validations.Elements<DataValidation>().ToList()) {
                    validation.Remove();
                    changed = true;
                }
            } else {
                var filter = ParseReferenceArgument(a1Range!);
                foreach (var validation in validations.Elements<DataValidation>().ToList()) {
                    string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                    if (!TryRemoveReferenceOverlap(range, filter, out var remaining)) {
                        continue;
                    }

                    if (remaining.Count == 0) {
                        validation.Remove();
                        changed = true;
                    } else {
                        string replacement = string.Join(" ", remaining);
                        validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = replacement };
                        changed = true;
                    }
                }
            }

            if (changed) {
                validations.Count = (uint)validations.Elements<DataValidation>().Count();
                WorksheetRoot.Save();
            }
        }

        private void InsertConditionalFormatting(ConditionalFormatting conditionalFormatting) {
            var worksheet = WorksheetRoot;
            var tableParts = worksheet.GetFirstChild<TableParts>();
            if (tableParts != null) {
                worksheet.InsertBefore(conditionalFormatting, tableParts);
                return;
            }

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                worksheet.InsertAfter(conditionalFormatting, autoFilter);
                return;
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                worksheet.InsertAfter(conditionalFormatting, sheetData);
            } else {
                worksheet.Append(conditionalFormatting);
            }
        }

        private int GetNextConditionalFormattingPriority() {
            if (_nextConditionalFormattingPriority > 0) {
                return _nextConditionalFormattingPriority++;
            }

            int priority = GetNextConditionalFormattingPriority(WorksheetRoot);
            _nextConditionalFormattingPriority = priority + 1;
            return priority;
        }

        private static int GetNextConditionalFormattingPriority(Worksheet worksheet) {
            int priority = 1;
            foreach (var conditional in worksheet.Elements<ConditionalFormatting>()) {
                foreach (var rule in conditional.Elements<ConditionalFormattingRule>()) {
                    int value = (int)(rule.Priority?.Value ?? 0);
                    if (value >= priority) {
                        priority = value + 1;
                    }
                }
            }

            return priority;
        }

        private static ReferenceListEnumerable SplitReferenceList(string referenceList) {
            return new ReferenceListEnumerable(referenceList);
        }

        private static (int r1, int c1, int r2, int c2) ParseReferenceArgument(string reference) {
            if (TryParseReference(reference, out var bounds)) {
                return bounds;
            }

            throw new ArgumentException($"Invalid A1 reference '{reference}'.");
        }

        private static bool ReferenceListOverlaps(string referenceList, (int r1, int c1, int r2, int c2) filter) {
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (TryParseReference(part, out var bounds) && RangesOverlapInclusive(filter, bounds)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryRemoveReferenceOverlap(string referenceList, (int r1, int c1, int r2, int c2) filter, out List<string> remaining) {
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (TryParseReference(part, out var bounds) && RangesOverlapInclusive(bounds, filter)) {
                    remaining = RemoveReferenceOverlapAfterOverlap(referenceList, filter);
                    return true;
                }
            }

            remaining = new List<string>(0);
            return false;
        }

        private static List<string> RemoveReferenceOverlapAfterOverlap(string referenceList, (int r1, int c1, int r2, int c2) filter) {
            var remaining = new List<string>();
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (!TryParseReference(part, out var bounds)) {
                    remaining.Add(part.ToString());
                    continue;
                }

                if (!RangesOverlapInclusive(bounds, filter)) {
                    remaining.Add(part.ToString());
                    continue;
                }

                AppendRangeDifference(remaining, bounds, filter);
            }

            return remaining;
        }

        private static void AppendRangeDifference(List<string> references, (int r1, int c1, int r2, int c2) source, (int r1, int c1, int r2, int c2) remove) {
            if (!RangesOverlapInclusive(source, remove)) {
                references.Add(ToReference(source.r1, source.c1, source.r2, source.c2));
                return;
            }

            int ir1 = Math.Max(source.r1, remove.r1);
            int ic1 = Math.Max(source.c1, remove.c1);
            int ir2 = Math.Min(source.r2, remove.r2);
            int ic2 = Math.Min(source.c2, remove.c2);

            if (source.r1 < ir1) {
                references.Add(ToReference(source.r1, source.c1, ir1 - 1, source.c2));
            }

            if (ir2 < source.r2) {
                references.Add(ToReference(ir2 + 1, source.c1, source.r2, source.c2));
            }

            if (source.c1 < ic1) {
                references.Add(ToReference(ir1, source.c1, ir2, ic1 - 1));
            }

            if (ic2 < source.c2) {
                references.Add(ToReference(ir1, ic2 + 1, ir2, source.c2));
            }
        }

        private readonly struct ReferenceListEnumerable {
            private readonly string _text;

            public ReferenceListEnumerable(string text) {
                _text = text;
            }

            public ReferenceListEnumerator GetEnumerator() {
                return new ReferenceListEnumerator(_text);
            }
        }

        private struct ReferenceListEnumerator {
            private readonly string _text;
            private int _index;

            public ReferenceListEnumerator(string text) {
                _text = text;
                _index = 0;
                Current = default;
            }

            public ReferenceListPart Current { get; private set; }

            public bool MoveNext() {
                int length = _text.Length;
                int start = _index;
                while (start < length && _text[start] == ' ') {
                    start++;
                }

                if (start >= length) {
                    _index = length;
                    Current = default;
                    return false;
                }

                int end = start + 1;
                while (end < length && _text[end] != ' ') {
                    end++;
                }

                _index = end + 1;
                Current = new ReferenceListPart(_text, start, end - start);
                return true;
            }
        }

        private readonly struct ReferenceListPart {
            public ReferenceListPart(string text, int start, int length) {
                Text = text;
                Start = start;
                Length = length;
            }

            public string Text { get; }

            public int Start { get; }

            public int Length { get; }

            public override string ToString() {
                return Start == 0 && Length == Text.Length ? Text : Text.Substring(Start, Length);
            }

            public void AppendTo(StringBuilder builder) {
                builder.Append(Text, Start, Length);
            }
        }
    }
}
