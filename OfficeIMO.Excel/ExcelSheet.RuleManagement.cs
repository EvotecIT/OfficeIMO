using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly List<string> EmptyReferenceList = new List<string>(0);

        /// <summary>
        /// Lists conditional formatting rules on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelConditionalFormattingInfo> GetConditionalFormattingRules(string? a1Range = null) {
            (int r1, int c1, int r2, int c2)? filter = string.IsNullOrWhiteSpace(a1Range) ? null : ParseReferenceArgument(a1Range!);
            var list = new List<ExcelConditionalFormattingInfo>();
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>()) {
                string range = conditional.SequenceOfReferences?.InnerText ?? string.Empty;
                if (filter.HasValue && !string.IsNullOrWhiteSpace(range)) {
                    if (!ReferenceListOverlaps(range, filter.Value)) continue;
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
            var filter = ParseReferenceArgument(a1Range);
            WriteLock(() => {
                var validations = WorksheetRoot.GetFirstChild<DataValidations>();
                if (validations == null) return;
                bool changed = false;
                bool showInputMessage = options.ShowInputMessage || !string.IsNullOrEmpty(options.Prompt) || !string.IsNullOrEmpty(options.PromptTitle);
                bool showErrorMessage = options.ShowErrorMessage || !string.IsNullOrEmpty(options.Error) || !string.IsNullOrEmpty(options.ErrorTitle);
                foreach (var validation in validations.Elements<DataValidation>()) {
                    string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
                    if (!ReferenceListOverlaps(range, filter)) continue;

                    bool validationChanged =
                        !string.Equals(validation.PromptTitle?.Value, options.PromptTitle, StringComparison.Ordinal)
                        || !string.Equals(validation.Prompt?.Value, options.Prompt, StringComparison.Ordinal)
                        || !string.Equals(validation.ErrorTitle?.Value, options.ErrorTitle, StringComparison.Ordinal)
                        || !string.Equals(validation.Error?.Value, options.Error, StringComparison.Ordinal)
                        || validation.ShowInputMessage?.Value != showInputMessage
                        || validation.ShowErrorMessage?.Value != showErrorMessage;

                    if (!validationChanged) continue;

                    validation.PromptTitle = options.PromptTitle;
                    validation.Prompt = options.Prompt;
                    validation.ErrorTitle = options.ErrorTitle;
                    validation.Error = options.Error;
                    validation.ShowInputMessage = showInputMessage;
                    validation.ShowErrorMessage = showErrorMessage;
                    changed = true;
                }

                if (changed) {
                    WorksheetRoot.Save();
                }
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

            remaining = EmptyReferenceList;
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
