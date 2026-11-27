using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxResolvedFormulaRangeCells = 100000;

        private bool TryResolveFormulaArgumentNumbers(string token, out List<double> numbers) {
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            return TryResolveFormulaArgumentNumbers(token, out numbers, ref remainingCellBudget);
        }

        private bool TryResolveFormulaArgumentNumbers(string token, out List<double> numbers, ref int remainingCellBudget) {
            numbers = new List<double>();
            if (TryResolveFormulaRange(token, out var values, ref remainingCellBudget, out bool isReference)) {
                foreach (var value in values) {
                    // Keep positional series numeric-only; dropping entries here would misalign paired ranges.
                    if (!value.IsNumericAggregateValue || !IsFinite(value.Number!.Value)) {
                        return false;
                    }

                    numbers.Add(value.Number!.Value);
                }

                return true;
            }
            if (isReference) return false;

            if (!TryEvaluateFormulaOrNumeric(token, out double numeric) || !IsFinite(numeric)) {
                return false;
            }

            numbers.Add(numeric);
            return true;
        }

        private bool TryEvaluateMultiCriteriaAggregate(string function, string args, out double result, out FormulaArgumentValue error) {
            result = 0;
            error = default;
            var tokens = SplitFormulaArguments(args);
            bool countOnly = function == "COUNTIFS";
            if (countOnly) {
                if (tokens.Count < 2 || tokens.Count % 2 != 0) {
                    return false;
                }
            } else if (tokens.Count < 3 || tokens.Count % 2 != 1) {
                return false;
            }

            int criteriaStart = countOnly ? 0 : 1;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            IReadOnlyList<FormulaArgumentValue>? aggregateValues = null;
            if (!countOnly && !TryResolveFormulaRange(tokens[0], out aggregateValues, ref remainingCellBudget)) {
                return false;
            }

            var criteriaSets = new List<(IReadOnlyList<FormulaArgumentValue> Values, FormulaCriteria Criteria)>();
            int expectedCount = aggregateValues?.Count ?? -1;
            for (int index = criteriaStart; index < tokens.Count; index += 2) {
                if (!TryResolveFormulaRange(tokens[index], out var criteriaValues, ref remainingCellBudget)
                    || !TryParseCriteria(tokens[index + 1], out var criteria)) {
                    return false;
                }

                if (expectedCount < 0) {
                    expectedCount = criteriaValues.Count;
                } else if (criteriaValues.Count != expectedCount && !_formulaReferenceError.IsError) {
                    return false;
                }

                criteriaSets.Add((criteriaValues, criteria));
            }
            if ((aggregateValues != null && ContainsUnresolvedFormula(aggregateValues))
                || criteriaSets.Any(set => ContainsUnresolvedFormula(set.Values))) return false;
            if (_formulaReferenceError.IsError) return true;

            var sets = criteriaSets.ToArray();
            if (countOnly || function == "SUMIFS") {
                // Count and sum in row order, as the list-based path below adds them, without collecting matches.
                int count = 0;
                double sum = 0;
                FormulaArgumentValue firstError = default;
                for (int rowIndex = 0; rowIndex < expectedCount; rowIndex++) {
                    if (!MatchesAllCriteria(sets, rowIndex)) continue;
                    count++;
                    if (countOnly) continue;
                    FormulaArgumentValue selected = aggregateValues![rowIndex];
                    if (selected.IsError) {
                        if (!firstError.IsError) firstError = selected;
                    } else if (selected.IsNumericAggregateValue) {
                        sum += selected.Number!.Value;
                    }
                }
                if (countOnly) {
                    result = count;
                    return true;
                }
                // Qualify every input before publishing a selected operand error.
                // Numeric filtering must not turn an error into an empty or partial aggregate.
                error = firstError;
                if (error.IsError) return false;
                result = sum;
                return true;
            }

            var matched = new List<FormulaArgumentValue>();
            for (int rowIndex = 0; rowIndex < expectedCount; rowIndex++) {
                if (MatchesAllCriteria(sets, rowIndex)) matched.Add(aggregateValues![rowIndex]);
            }

            error = matched.FirstOrDefault(value => value.IsError);
            if (error.IsError) return false;

            var numbers = matched.Where(value => value.IsNumericAggregateValue).Select(value => value.Number!.Value).ToList();

            if (numbers.Count == 0) {
                if (function == "AVERAGEIFS") error = FormulaArgumentValue.Error("#DIV/0!");
                else if (matched.Count == 0 && (function == "MINIFS" || function == "MAXIFS")) return true;
                return false;
            }

            if (function == "MINIFS") {
                result = numbers.Min();
            } else if (function == "MAXIFS") {
                result = numbers.Max();
            } else {
                result = numbers.Average();
            }
            return true;
        }

        private static bool MatchesAllCriteria((IReadOnlyList<FormulaArgumentValue> Values, FormulaCriteria Criteria)[] sets, int rowIndex) {
            for (int index = 0; index < sets.Length; index++) {
                if (!MatchesCriteria(sets[index].Values[rowIndex], sets[index].Criteria)) return false;
            }
            return true;
        }

        private static bool ContainsUnresolvedFormula(IReadOnlyList<FormulaArgumentValue> values) {
            for (int index = 0; index < values.Count; index++) {
                if (values[index].IsUnresolvedFormula) return true;
            }
            return false;
        }

        private bool TryEvaluateCondition(string condition, out bool result) {
            return TryEvaluateCondition(condition, out result, out FormulaArgumentValue error) && !error.IsError;
        }

        private bool TryEvaluateCondition(string condition, out bool result, out FormulaArgumentValue error) {
            result = false;
            error = default;
            if (!TryResolveFormulaArgument(condition, out FormulaArgumentValue value) || value.IsUnresolvedFormula) return false;
            if (value.IsError) { error = value; return true; }
            if (!value.HasValue && value.SourceCellKind == ExcelCellDataKind.Blank) return true;
            if (!value.Number.HasValue) return false;
            result = Math.Abs(value.Number.Value) >= double.Epsilon;
            return true;
        }

        private static bool TryCompareFormulaValues(FormulaArgumentValue left, string comparisonOperator, FormulaArgumentValue right, out bool result) {
            result = false;
            // A blank reference compares as zero against a numeric operand;
            // an actual empty text value retains text comparison semantics.
            if (!left.HasValue && left.SourceCellKind == ExcelCellDataKind.Blank && right.Number.HasValue)
                left = new FormulaArgumentValue(0d, "0");
            if (!right.HasValue && right.SourceCellKind == ExcelCellDataKind.Blank && left.Number.HasValue)
                right = new FormulaArgumentValue(0d, "0");
            int comparison;
            if (left.Number.HasValue && right.Number.HasValue) {
                comparison = CompareFormulaNumbers(left.Number.Value, right.Number.Value);
            } else {
                string leftText = left.Text ?? (left.Number.HasValue ? InvariantNumberText.Get(left.Number.Value) : string.Empty);
                string rightText = right.Text ?? (right.Number.HasValue ? InvariantNumberText.Get(right.Number.Value) : string.Empty);
                comparison = string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase);
            }

            switch (comparisonOperator) {
                case ">":
                    result = comparison > 0;
                    return true;
                case "<":
                    result = comparison < 0;
                    return true;
                case ">=":
                    result = comparison >= 0;
                    return true;
                case "<=":
                    result = comparison <= 0;
                    return true;
                case "=":
                    result = comparison == 0;
                    return true;
                case "<>":
                    result = comparison != 0;
                    return true;
                default:
                    return false;
            }
        }

        private bool TryParseCriteria(string token, out FormulaCriteria criteria) {
            string value = token.Trim();

            if (!TryResolveFormulaArgument(value, out FormulaArgumentValue criteriaValue)
                || criteriaValue.IsUnresolvedFormula) {
                criteria = default;
                return false;
            }
            if (criteriaValue.IsError) {
                criteria = new FormulaCriteria("=", string.Empty, null, criteriaValue.ErrorCode);
                return true;
            }
            value = FormulaValueToText(NormalizeLookupResult(criteriaValue));

            string op = "=";
            foreach (string candidate in new[] { ">=", "<=", "<>", ">", "<", "=" }) {
                if (value.StartsWith(candidate, StringComparison.Ordinal)) {
                    op = candidate;
                    value = value.Substring(candidate.Length);
                    break;
                }
            }

            value = value.Trim();
            double? number = OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)
                ? numeric
                : (double?)null;

            criteria = new FormulaCriteria(op, value, number);
            return value.Length > 0 || op == "=" || op == "<>";
        }

        private static bool MatchesCriteria(in FormulaArgumentValue value, in FormulaCriteria criteria) {
            if (criteria.ErrorCode != null) {
                return value.IsError && string.Equals(value.ErrorCode, criteria.ErrorCode, StringComparison.Ordinal);
            }
            // Error cells are distinct from ordinary text, including a literal "#N/A".
            // They satisfy an ordinary not-equal criterion, but no text/numeric ordering or wildcard.
            if (value.IsError) return criteria.Kind == FormulaCriteriaOperator.NotEqual;

            if (criteria.Number.HasValue && value.Number.HasValue) {
                double left = value.Number.Value;
                double right = criteria.Number.Value;
                switch (criteria.Kind) {
                    case FormulaCriteriaOperator.GreaterThan:
                        return left > right;
                    case FormulaCriteriaOperator.LessThan:
                        return left < right;
                    case FormulaCriteriaOperator.GreaterThanOrEqual:
                        return left >= right;
                    case FormulaCriteriaOperator.LessThanOrEqual:
                        return left <= right;
                    case FormulaCriteriaOperator.NotEqual:
                        return left != right;
                    default:
                        return left == right;
                }
            }

            string text = value.Text ?? string.Empty;
            switch (criteria.Kind) {
                case FormulaCriteriaOperator.GreaterThan:
                    return string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase) > 0;
                case FormulaCriteriaOperator.LessThan:
                    return string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase) < 0;
                case FormulaCriteriaOperator.GreaterThanOrEqual:
                    return string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase) >= 0;
                case FormulaCriteriaOperator.LessThanOrEqual:
                    return string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase) <= 0;
                case FormulaCriteriaOperator.NotEqual:
                    return !MatchesTextCriteria(text, criteria);
                default:
                    return MatchesTextCriteria(text, criteria);
            }
        }

        private static bool MatchesTextCriteria(string text, in FormulaCriteria criteria) =>
            criteria.IsTextPattern
                ? MatchesFormulaWildcard(criteria.Text, text)
                : string.Equals(text, criteria.Text, StringComparison.OrdinalIgnoreCase);

        private static bool IsFormulaBlankValue(FormulaArgumentValue value) {
            return !value.IsUnresolvedFormula && !value.Number.HasValue && string.IsNullOrEmpty(value.Text);
        }

        private static bool TryGetSupportedDecimalPlaces(double value, out int digits) {
            digits = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - digits) < 0.0000001 && digits >= -15 && digits <= 15;
        }

        private static double RoundAtDigits(double value, int digits, MidpointRounding mode) {
            // Excel rounds decimal midpoints despite the binary representation of values
            // such as 1.255. Decimal conversion normalizes the input's 15 significant digits.
            if (Math.Abs(value) < (double)decimal.MaxValue) {
                try {
                    decimal number = OfficeDecimalConversion.FromDouble(value);
                    if (digits >= 0) return (double)decimal.Round(number, digits, mode);
                    decimal scale = OfficeDecimalConversion.FromDouble(Math.Pow(10, -digits));
                    return (double)(decimal.Round(number / scale, 0, mode) * scale);
                } catch (OverflowException) {
                    // Rounding can cross decimal's upper bound; doubles still represent it.
                }
            }
            if (digits >= 0) {
                return Math.Round(value, digits, mode);
            }

            double factor = Math.Pow(10, -digits);
            return Math.Round(value / factor, 0, mode) * factor;
        }

        private static double CalculateMedian(List<double> numbers) {
            var sorted = numbers.OrderBy(value => value).ToList();
            int middle = sorted.Count / 2;
            if (sorted.Count % 2 == 1) return sorted[middle];
            double left = sorted[middle - 1], right = sorted[middle];
            return left >= 0d || right <= 0d ? left + (right - left) / 2d : (left + right) / 2d;
        }

        private static int CompareFormulaNumbers(double left, double right) {
            // Comparison follows spreadsheet decimal precision without changing the
            // full-precision numbers stored in caches or passed to financial solvers.
            double normalizedLeft = NormalizeFormulaComparisonNumber(left);
            double normalizedRight = NormalizeFormulaComparisonNumber(right);
            return normalizedLeft.CompareTo(normalizedRight);
        }

        private static double NormalizeFormulaComparisonNumber(double value) {
            // A whole number below 10^15 has at most 15 significant digits, so the G15 round trip returns it unchanged.
            if (Math.Abs(value) < 1e15 && Math.Floor(value) == value) return value;
            double normalized;
#if NETCOREAPP
            Span<char> text = stackalloc char[32];
            if (value.TryFormat(text, out int written, "G15", CultureInfo.InvariantCulture))
                return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text.Slice(0, written), NumberStyles.Float, CultureInfo.InvariantCulture, out normalized)
                    && !double.IsInfinity(normalized) ? normalized : value;
#endif
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out normalized)
                && !double.IsInfinity(normalized) ? normalized : value;
        }

        private static bool TryGetWholeNumber(double value, out int number) {
            number = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - number) < 0.0000001;
        }

        private bool TryGetWholeNumberArgument(string token, out int number) {
            number = 0;
            return TryEvaluateFormulaOrNumeric(token, out double value)
                && TryGetWholeNumber(value, out number);
        }

        private double ToExcelDateSerial(DateTime value) {
            return ExcelDateSystemConverter.ToSerial(value, _excelDocument.DateSystem);
        }

        private bool TryGetDateFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = ExcelDateSystemConverter.FromSerial(value, _excelDocument.DateSystem).Date;
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private bool TryGetDateTimeFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = ExcelDateSystemConverter.FromSerial(value, _excelDocument.DateSystem);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

    }
}
