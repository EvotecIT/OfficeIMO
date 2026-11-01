using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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
            List<FormulaArgumentValue>? aggregateValues = null;
            if (!countOnly && !TryResolveFormulaRange(tokens[0], out aggregateValues, ref remainingCellBudget)) {
                return false;
            }

            var criteriaSets = new List<(List<FormulaArgumentValue> Values, FormulaCriteria Criteria)>();
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
            if ((aggregateValues?.Any(value => value.IsUnresolvedFormula) ?? false)
                || criteriaSets.Any(set => set.Values.Any(value => value.IsUnresolvedFormula))) return false;
            if (_formulaReferenceError.IsError) return true;

            var matched = new List<FormulaArgumentValue>();
            for (int rowIndex = 0; rowIndex < expectedCount; rowIndex++) {
                bool matches = true;
                foreach (var criteriaSet in criteriaSets) {
                    if (!MatchesCriteria(criteriaSet.Values[rowIndex], criteriaSet.Criteria)) {
                        matches = false;
                        break;
                    }
                }

                if (matches) {
                    matched.Add(countOnly ? criteriaSets[0].Values[rowIndex] : aggregateValues![rowIndex]);
                }
            }

            if (countOnly) {
                result = matched.Count;
                return true;
            }

            // Qualify every input before publishing a selected operand error.
            // Numeric filtering must not turn an error into an empty or partial aggregate.
            error = matched.FirstOrDefault(value => value.IsError);
            if (error.IsError) return false;

            var numbers = matched.Where(value => value.IsNumericAggregateValue).Select(value => value.Number!.Value).ToList();
            if (function == "SUMIFS") {
                result = numbers.Sum();
                return true;
            }

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
                comparison = left.Number.Value.CompareTo(right.Number.Value);
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
            double? number = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)
                ? numeric
                : (double?)null;

            criteria = new FormulaCriteria(op, value, number);
            return value.Length > 0 || op == "=" || op == "<>";
        }

        private static bool MatchesCriteria(FormulaArgumentValue value, FormulaCriteria criteria) {
            if (criteria.ErrorCode != null) {
                return value.IsError && string.Equals(value.ErrorCode, criteria.ErrorCode, StringComparison.Ordinal);
            }
            // Error cells are distinct from ordinary text, including a literal "#N/A".
            // They satisfy an ordinary not-equal criterion, but no text/numeric ordering or wildcard.
            if (value.IsError) return criteria.Operator == "<>";

            if (criteria.Number.HasValue && value.Number.HasValue) {
                double left = value.Number.Value;
                double right = criteria.Number.Value;
                switch (criteria.Operator) {
                    case ">":
                        return left > right;
                    case "<":
                        return left < right;
                    case ">=":
                        return left >= right;
                    case "<=":
                        return left <= right;
                    case "<>":
                        return left != right;
                    default:
                        return left == right;
                }
            }

            string text = value.Text ?? string.Empty;
            int comparison = string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase);
            switch (criteria.Operator) {
                case ">":
                    return comparison > 0;
                case "<":
                    return comparison < 0;
                case ">=":
                    return comparison >= 0;
                case "<=":
                    return comparison <= 0;
                case "<>":
                    return !MatchesTextCriteria(text, criteria.Text);
                default:
                    return MatchesTextCriteria(text, criteria.Text);
            }
        }

        private static bool MatchesTextCriteria(string text, string criteria) {
            if (criteria.IndexOf('*') < 0 && criteria.IndexOf('?') < 0 && criteria.IndexOf('~') < 0) {
                return string.Equals(text, criteria, StringComparison.OrdinalIgnoreCase);
            }

            string pattern = "\\A" + CreateFormulaWildcardPattern(criteria) + "\\z";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, FormulaRegexTimeout);
        }

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
                    decimal number = (decimal)value;
                    if (digits >= 0) return (double)decimal.Round(number, digits, mode);
                    decimal scale = (decimal)Math.Pow(10, -digits);
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
            return sorted.Count % 2 == 1
                ? sorted[middle]
                : (sorted[middle - 1] + sorted[middle]) / 2d;
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
