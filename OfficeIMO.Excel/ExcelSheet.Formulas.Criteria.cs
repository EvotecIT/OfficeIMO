using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxResolvedFormulaRangeCells = 100000;

        private bool TryResolveFormulaArgumentNumbers(string token, out List<double> numbers) {
            numbers = new List<double>();
            if (TryResolveFormulaRange(token, out var values)) {
                foreach (var value in values) {
                    if (!value.Number.HasValue) {
                        return false;
                    }

                    numbers.Add(value.Number.Value);
                }

                return true;
            }

            if (!TryEvaluateFormulaOrNumeric(token, out double numeric)) {
                return false;
            }

            numbers.Add(numeric);
            return true;
        }

        private bool TryEvaluateMultiCriteriaAggregate(string function, string args, out double result) {
            result = 0;
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
            List<FormulaArgumentValue>? aggregateValues = null;
            if (!countOnly && !TryResolveFormulaRange(tokens[0], out aggregateValues)) {
                return false;
            }

            var criteriaSets = new List<(List<FormulaArgumentValue> Values, FormulaCriteria Criteria)>();
            int expectedCount = aggregateValues?.Count ?? -1;
            for (int index = criteriaStart; index < tokens.Count; index += 2) {
                if (!TryResolveFormulaRange(tokens[index], out var criteriaValues)
                    || !TryParseCriteria(tokens[index + 1], out var criteria)) {
                    return false;
                }

                if (expectedCount < 0) {
                    expectedCount = criteriaValues.Count;
                } else if (criteriaValues.Count != expectedCount) {
                    return false;
                }

                criteriaSets.Add((criteriaValues, criteria));
            }

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

            var numbers = matched.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "SUMIFS") {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
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
            result = false;
            if (TryResolveBooleanArgument(condition, out bool booleanResult)) {
                result = booleanResult;
                return true;
            }

            if (TrySplitFormulaComparison(condition, out string leftToken, out string comparisonOperator, out string rightToken)
                && TryResolveFormulaArgument(leftToken, out FormulaArgumentValue leftValue)
                && TryResolveFormulaArgument(rightToken, out FormulaArgumentValue rightValue)
                && !leftValue.IsUnresolvedFormula
                && !rightValue.IsUnresolvedFormula
                && leftValue.HasValue
                && rightValue.HasValue) {
                return TryCompareFormulaValues(leftValue, comparisonOperator, rightValue, out result);
            }

            var comparison = SimpleComparisonFormulaRegex.Match(condition);
            if (comparison.Success) {
                if (!TryResolveNumericOperand(comparison.Groups[1].Value, out double left)
                    || !TryResolveNumericOperand(comparison.Groups[3].Value, out double right)) {
                    return false;
                }

                switch (comparison.Groups[2].Value) {
                    case ">":
                        result = left > right;
                        return true;
                    case "<":
                        result = left < right;
                        return true;
                    case ">=":
                        result = left >= right;
                        return true;
                    case "<=":
                        result = left <= right;
                        return true;
                    case "=":
                        result = Math.Abs(left - right) < 0.0000001;
                        return true;
                    case "<>":
                        result = Math.Abs(left - right) >= 0.0000001;
                        return true;
                }
            }

            if (TryEvaluateFormula(condition, out double formulaValue)) {
                result = Math.Abs(formulaValue) >= double.Epsilon;
                return true;
            }

            if (!TryResolveNumericOperand(condition, out double value)) {
                return false;
            }

            result = Math.Abs(value) >= double.Epsilon;
            return true;
        }

        private static bool TrySplitFormulaComparison(string condition, out string left, out string comparisonOperator, out string right) {
            left = string.Empty;
            comparisonOperator = string.Empty;
            right = string.Empty;
            string value = condition.Trim();
            int depth = 0;
            int bracketDepth = 0;
            bool inString = false;

            for (int index = 0; index < value.Length; index++) {
                char ch = value[index];
                if (ch == '"') {
                    if (inString && index + 1 < value.Length && value[index + 1] == '"') {
                        index++;
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (inString) {
                    continue;
                }

                if (ch == '(') {
                    depth++;
                    continue;
                }

                if (ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }

                    continue;
                }

                if (ch == '[') {
                    bracketDepth++;
                    continue;
                }

                if (ch == ']') {
                    bracketDepth--;
                    if (bracketDepth < 0) {
                        return false;
                    }

                    continue;
                }

                if (depth != 0 || bracketDepth != 0) {
                    continue;
                }

                string? op = null;
                if (index + 1 < value.Length) {
                    string pair = value.Substring(index, 2);
                    if (pair == ">=" || pair == "<=" || pair == "<>") {
                        op = pair;
                    }
                }

                if (op == null && (ch == '=' || ch == '>' || ch == '<')) {
                    op = ch.ToString();
                }

                if (op == null) {
                    continue;
                }

                left = value.Substring(0, index).Trim();
                comparisonOperator = op;
                right = value.Substring(index + op.Length).Trim();
                return left.Length > 0 && right.Length > 0;
            }

            return false;
        }

        private static bool TryCompareFormulaValues(FormulaArgumentValue left, string comparisonOperator, FormulaArgumentValue right, out bool result) {
            result = false;
            int comparison;
            if (left.Number.HasValue && right.Number.HasValue) {
                double delta = left.Number.Value - right.Number.Value;
                comparison = Math.Abs(delta) < 0.0000001 ? 0 : delta < 0 ? -1 : 1;
            } else {
                string leftText = left.Text ?? left.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                string rightText = right.Text ?? right.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
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

        private bool TryResolveFormulaRange(string token, out List<FormulaArgumentValue> values) {
            values = new List<FormulaArgumentValue>();
            if (!TryResolveFormulaRangeReference(token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            long rowCount = (long)r2 - r1 + 1;
            long columnCount = (long)c2 - c1 + 1;
            long cellCount = rowCount * columnCount;
            if (cellCount > MaxResolvedFormulaRangeCells) {
                return false;
            }

            values = new List<FormulaArgumentValue>((int)cellCount);
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    values.Add(sheet.ResolveCellArgument(row, column));
                }
            }

            return true;
        }

        private bool TryParseCriteria(string token, out FormulaCriteria criteria) {
            string value = token.Trim();

            if (TryResolveFormulaArgument(value, out FormulaArgumentValue criteriaValue) && !criteriaValue.IsUnresolvedFormula && criteriaValue.HasValue) {
                value = FormulaValueToText(criteriaValue);
            } else if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"') {
                value = value.Substring(1, value.Length - 2);
            }

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
            return value.Length > 0 || op == "=";
        }

        private static bool MatchesCriteria(FormulaArgumentValue value, FormulaCriteria criteria) {
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
                        return Math.Abs(left - right) >= 0.0000001;
                    default:
                        return Math.Abs(left - right) < 0.0000001;
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
            if (criteria.IndexOf('*') < 0 && criteria.IndexOf('?') < 0) {
                return string.Equals(text, criteria, StringComparison.OrdinalIgnoreCase);
            }

            string pattern = "^" + Regex.Escape(criteria).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase, FormulaRegexTimeout);
        }

        private static bool IsFormulaBlankValue(FormulaArgumentValue value) {
            return !value.IsUnresolvedFormula && !value.Number.HasValue && string.IsNullOrEmpty(value.Text);
        }

        private static bool TryGetSupportedDecimalPlaces(double value, out int digits) {
            digits = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - digits) < 0.0000001 && digits >= -15 && digits <= 15;
        }

        private static double RoundAtDigits(double value, int digits, MidpointRounding mode) {
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

        private static bool TryGetDateFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = DateTime.FromOADate(value).Date;
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static bool TryGetDateTimeFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = DateTime.FromOADate(value);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

    }
}
