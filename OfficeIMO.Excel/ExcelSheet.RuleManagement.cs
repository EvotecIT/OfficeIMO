using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Lists conditional formatting rules on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelConditionalFormattingInfo> GetConditionalFormattingRules(string? a1Range = null) {
            (int r1, int c1, int r2, int c2)? filter = string.IsNullOrWhiteSpace(a1Range) ? null : A1.ParseRange(a1Range!);
            var list = new List<ExcelConditionalFormattingInfo>();
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>()) {
                string range = conditional.SequenceOfReferences?.InnerText ?? string.Empty;
                if (filter.HasValue && !string.IsNullOrWhiteSpace(range)) {
                    bool overlaps = SplitReferenceList(range)
                        .Any(part => RangesOverlapInclusive(filter.Value, part.IndexOf(':') >= 0 ? A1.ParseRange(part) : CellAsRange(part)));
                    if (!overlaps) continue;
                }

                foreach (var rule in conditional.Elements<ConditionalFormattingRule>()) {
                    list.Add(new ExcelConditionalFormattingInfo {
                        Range = range,
                        Type = rule.Type?.Value.ToString() ?? string.Empty,
                        Operator = rule.Operator?.Value.ToString(),
                        Priority = (int)(rule.Priority?.Value ?? 0),
                        StopIfTrue = rule.StopIfTrue?.Value ?? false,
                        Formulas = rule.Elements<Formula>().Select(f => f.Text ?? string.Empty).ToArray()
                    });
                }
            }

            return list;
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
        public void AddConditionalFormulaRule(string range, string formula, bool stopIfTrue = false) {
            AddConditionalRuleCore(range, ConditionalFormatValues.Expression, null, new[] { formula }, stopIfTrue);
        }

        /// <summary>
        /// Adds a duplicate-values conditional formatting rule.
        /// </summary>
        public void AddConditionalDuplicateValuesRule(string range) {
            AddConditionalRuleCore(range, ConditionalFormatValues.DuplicateValues, null, Array.Empty<string>(), stopIfTrue: false);
        }

        /// <summary>
        /// Adds a top/bottom conditional formatting rule.
        /// </summary>
        public void AddConditionalTopBottomRule(string range, uint rank, bool bottom = false, bool percent = false) {
            if (rank == 0) throw new ArgumentOutOfRangeException(nameof(rank));
            AddConditionalRuleCore(range, ConditionalFormatValues.Top10, rule => {
                rule.Rank = rank;
                rule.Bottom = bottom;
                rule.Percent = percent;
            }, Array.Empty<string>(), stopIfTrue: false);
        }

        /// <summary>
        /// Lists data validation rules on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelDataValidationInfo> GetDataValidations(string? a1Range = null) {
            (int r1, int c1, int r2, int c2)? filter = string.IsNullOrWhiteSpace(a1Range) ? null : A1.ParseRange(a1Range!);
            var result = new List<ExcelDataValidationInfo>();
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) return result;

            foreach (var validation in validations.Elements<DataValidation>()) {
                string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                if (filter.HasValue) {
                    bool overlaps = SplitReferenceList(range)
                        .Any(part => RangesOverlapInclusive(filter.Value, part.IndexOf(':') >= 0 ? A1.ParseRange(part) : CellAsRange(part)));
                    if (!overlaps) continue;
                }

                result.Add(new ExcelDataValidationInfo {
                    Range = range,
                    Type = validation.Type?.Value.ToString() ?? string.Empty,
                    Operator = validation.Operator?.Value.ToString(),
                    AllowBlank = validation.AllowBlank?.Value ?? false,
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
            var filter = A1.ParseRange(a1Range);
            WriteLock(() => {
                var validations = WorksheetRoot.GetFirstChild<DataValidations>();
                if (validations == null) return;
                foreach (var validation in validations.Elements<DataValidation>()) {
                    string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                    bool overlaps = SplitReferenceList(range)
                        .Any(part => RangesOverlapInclusive(filter, part.IndexOf(':') >= 0 ? A1.ParseRange(part) : CellAsRange(part)));
                    if (!overlaps) continue;

                    validation.PromptTitle = options.PromptTitle;
                    validation.Prompt = options.Prompt;
                    validation.ErrorTitle = options.ErrorTitle;
                    validation.Error = options.Error;
                    validation.ShowInputMessage = options.ShowInputMessage || !string.IsNullOrEmpty(options.Prompt) || !string.IsNullOrEmpty(options.PromptTitle);
                    validation.ShowErrorMessage = options.ShowErrorMessage || !string.IsNullOrEmpty(options.Error) || !string.IsNullOrEmpty(options.ErrorTitle);
                }

                WorksheetRoot.Save();
            });
        }

        private void AddConditionalRuleCore(string range, ConditionalFormatValues type, Action<ConditionalFormattingRule>? configure, IReadOnlyList<string> formulas, bool stopIfTrue) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            WriteLock(() => {
                int priority = WorksheetRoot.Descendants<ConditionalFormattingRule>().Select(r => r.Priority?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
                var conditional = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };
                var rule = new ConditionalFormattingRule {
                    Type = type,
                    Priority = priority,
                    StopIfTrue = stopIfTrue
                };
                configure?.Invoke(rule);
                foreach (var formula in formulas) {
                    rule.Append(new Formula(formula));
                }
                conditional.Append(rule);
                InsertConditionalFormatting(conditional);
                WorksheetRoot.Save();
            });
        }

        private void ClearConditionalFormattingCore(string? a1Range) {
            if (string.IsNullOrWhiteSpace(a1Range)) {
                foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                    conditional.Remove();
                }
            } else {
                var filter = A1.ParseRange(a1Range!);
                foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                    string range = conditional.SequenceOfReferences?.InnerText ?? string.Empty;
                    var remaining = RemoveReferenceOverlap(range, filter);
                    if (remaining.Count == 0) {
                        conditional.Remove();
                    } else if (!string.Equals(range, string.Join(" ", remaining), StringComparison.OrdinalIgnoreCase)) {
                        conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", remaining) };
                    }
                }
            }

            WorksheetRoot.Save();
        }

        private void RemoveDataValidationsCore(string? a1Range) {
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) return;

            if (string.IsNullOrWhiteSpace(a1Range)) {
                validations.RemoveAllChildren<DataValidation>();
            } else {
                var filter = A1.ParseRange(a1Range!);
                foreach (var validation in validations.Elements<DataValidation>().ToList()) {
                    string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                    var remaining = RemoveReferenceOverlap(range, filter);
                    if (remaining.Count == 0) {
                        validation.Remove();
                    } else if (!string.Equals(range, string.Join(" ", remaining), StringComparison.OrdinalIgnoreCase)) {
                        validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", remaining) };
                    }
                }
            }

            validations.Count = (uint)validations.Elements<DataValidation>().Count();
            WorksheetRoot.Save();
        }

        private void InsertConditionalFormatting(ConditionalFormatting conditionalFormatting) {
            var worksheet = WorksheetRoot;
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            if (tableParts != null) {
                worksheet.InsertBefore(conditionalFormatting, tableParts);
                return;
            }

            var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
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

        private static string[] SplitReferenceList(string referenceList) {
            return referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static List<string> RemoveReferenceOverlap(string referenceList, (int r1, int c1, int r2, int c2) filter) {
            var remaining = new List<string>();
            foreach (string part in SplitReferenceList(referenceList)) {
                if (!TryParseReference(part, out var bounds)) {
                    remaining.Add(part);
                    continue;
                }

                foreach (var segment in SubtractRange(bounds, filter)) {
                    remaining.Add(ToReference(segment.r1, segment.c1, segment.r2, segment.c2));
                }
            }

            return remaining;
        }

        private static IEnumerable<(int r1, int c1, int r2, int c2)> SubtractRange((int r1, int c1, int r2, int c2) source, (int r1, int c1, int r2, int c2) remove) {
            if (!RangesOverlapInclusive(source, remove)) {
                yield return source;
                yield break;
            }

            int ir1 = Math.Max(source.r1, remove.r1);
            int ic1 = Math.Max(source.c1, remove.c1);
            int ir2 = Math.Min(source.r2, remove.r2);
            int ic2 = Math.Min(source.c2, remove.c2);

            if (source.r1 < ir1) {
                yield return (source.r1, source.c1, ir1 - 1, source.c2);
            }

            if (ir2 < source.r2) {
                yield return (ir2 + 1, source.c1, source.r2, source.c2);
            }

            if (source.c1 < ic1) {
                yield return (ir1, source.c1, ir2, ic1 - 1);
            }

            if (ic2 < source.c2) {
                yield return (ir1, ic2 + 1, ir2, source.c2);
            }
        }
    }
}
