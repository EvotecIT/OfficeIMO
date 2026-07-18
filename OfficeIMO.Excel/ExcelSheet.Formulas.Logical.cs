using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateErrorFallback(string function, string args, out double result) {
            result = 0;
            bool success = function == "IFNA"
                ? TryEvaluateIfNaValue(args, out FormulaArgumentValue value)
                : TryEvaluateIfErrorValue(args, out value);
            if (!success || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateIf(string args, out double result) {
            result = 0;
            if (!TryEvaluateIfValue(args, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateIfErrorValue(string args, out FormulaArgumentValue result) {
            return TryEvaluateErrorFallbackValue(args, _ => true, out result);
        }

        private bool TryEvaluateIfNaValue(string args, out FormulaArgumentValue result) {
            return TryEvaluateErrorFallbackValue(args, errorCode => string.Equals(errorCode, "#N/A", StringComparison.OrdinalIgnoreCase), out result);
        }

        private bool TryEvaluateErrorFallbackValue(string args, Func<string, bool> shouldUseFallback, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 2) {
                return false;
            }

            if (!TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue candidate)
                || candidate.IsUnresolvedFormula
                || !candidate.HasValue) {
                if (!TryInferFormulaErrorArgument(tokens[0], out string inferredErrorCode)) {
                    return false;
                }

                candidate = FormulaArgumentValue.Error(inferredErrorCode);
            }

            if (!candidate.IsError || !shouldUseFallback(candidate.ErrorCode ?? "#VALUE!")) {
                result = candidate;
                return true;
            }

            return TryResolveFormulaArgument(tokens[1], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateIfValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 3
                || !TryEvaluateCondition(tokens[0], out bool condition)
                || !TryResolveFormulaArgument(tokens[condition ? 1 : 2], out result)
                || result.IsUnresolvedFormula
                || !result.HasValue) {
                return false;
            }

            return true;
        }

        private bool TryEvaluateIfsValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || tokens.Count % 2 != 0) {
                return false;
            }

            for (int index = 0; index < tokens.Count; index += 2) {
                if (!TryEvaluateCondition(tokens[index], out bool condition)) {
                    return false;
                }

                if (!condition) {
                    continue;
                }

                return TryResolveFormulaArgument(tokens[index + 1], out result)
                    && !result.IsUnresolvedFormula
                    && result.HasValue;
            }

            return false;
        }

        private bool TryEvaluateSwitchValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 3) {
                return false;
            }

            if (!TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue expression)
                || expression.IsUnresolvedFormula
                || !expression.HasValue) {
                return false;
            }

            int pairEnd = tokens.Count;
            bool hasDefault = tokens.Count % 2 == 0;
            if (hasDefault) {
                pairEnd--;
            }

            for (int index = 1; index < pairEnd; index += 2) {
                if (!TryResolveFormulaArgument(tokens[index], out FormulaArgumentValue candidate)
                    || candidate.IsUnresolvedFormula
                    || !candidate.HasValue) {
                    return false;
                }

                if (!FormulaValuesEqual(expression, candidate)) {
                    continue;
                }

                return TryResolveFormulaArgument(tokens[index + 1], out result)
                    && !result.IsUnresolvedFormula
                    && result.HasValue;
            }

            return hasDefault
                && TryResolveFormulaArgument(tokens[tokens.Count - 1], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateChooseValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || !TryGetWholeNumberArgument(tokens[0], out int index)) {
                return false;
            }

            if (index < 1 || index >= tokens.Count) {
                return false;
            }

            return TryResolveFormulaArgument(tokens[index], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateInfoFunction(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "ISFORMULA") {
                if (tokens.Count != 1 || !TryResolveFormulaReferenceArgument(tokens[0], out ExcelSheet referenceSheet, out int referenceRow, out int referenceColumn)) {
                    return false;
                }

                bool isFormula = referenceSheet.TryGetExistingCell(referenceRow, referenceColumn)?.CellFormula != null;
                result = new FormulaArgumentValue(isFormula ? 1d : 0d, isFormula ? "1" : "0");
                return true;
            }

            if (tokens.Count != 1
                || !TryResolveInfoArgument(tokens[0], out FormulaArgumentValue value)
                || value.IsUnresolvedFormula) {
                return false;
            }

            bool matches;
            switch (function) {
                case "ISBLANK":
                    matches = !value.HasValue;
                    break;
                case "ISNUMBER":
                    matches = value.Number.HasValue;
                    break;
                case "ISTEXT":
                    matches = value.Text != null && !value.Number.HasValue;
                    break;
                case "ISERROR":
                    matches = value.IsError;
                    break;
                case "ISERR":
                    matches = value.IsError && !string.Equals(value.ErrorCode, "#N/A", StringComparison.OrdinalIgnoreCase);
                    break;
                case "ISNA":
                    matches = value.IsError && string.Equals(value.ErrorCode, "#N/A", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    return false;
            }

            result = new FormulaArgumentValue(matches ? 1d : 0d, matches ? "1" : "0");
            return true;
        }

        private bool TryEvaluateReferenceShapeFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1) {
                return false;
            }

            if (function == "ROW" || function == "COLUMN") {
                if (!TryResolveFormulaReferenceArgument(tokens[0], out _, out int row, out int column)) {
                    return false;
                }

                result = function == "ROW" ? row : column;
                return true;
            }

            if (!TryResolveFormulaRangeReference(
                    tokens[0],
                    out _,
                    out int r1,
                    out int c1,
                    out int r2,
                    out int c2)
                && !TryParseQualifiedFormulaWholeRange(
                    tokens[0],
                    null,
                    out _,
                    out r1,
                    out c1,
                    out r2,
                    out c2,
                    out _)) {
                return false;
            }

            result = function == "ROWS"
                ? r2 - r1 + 1
                : c2 - c1 + 1;
            return result > 0;
        }

        private bool TryGetFormulaTextArgument(string token, out string formulaText) {
            formulaText = string.Empty;
            if (!TryResolveFormulaReferenceArgument(token, out ExcelSheet referenceSheet, out int row, out int column)) {
                return false;
            }

            Cell? cell = referenceSheet.TryGetExistingCell(row, column);
            formulaText = cell?.CellFormula == null
                ? string.Empty
                : referenceSheet.ResolveCellFormulaText(
                    cell,
                    cell.CellFormula.FormulaType?.Value == CellFormulaValues.Shared
                        && string.IsNullOrEmpty(cell.CellFormula.Text)
                        ? GetFormulaEvaluationSharedDefinitions(referenceSheet)
                        : null);
            return formulaText.Length > 0;
        }

        private bool TryResolveInfoArgument(string token, out FormulaArgumentValue value) {
            if (TryResolveFormulaArgument(token, out value)) {
                return true;
            }

            if (TryInferFormulaErrorArgument(token, out string errorCode)) {
                value = FormulaArgumentValue.Error(errorCode);
                return true;
            }

            value = default;
            return false;
        }

        private bool TryInferFormulaErrorArgument(string token, out string errorCode) {
            string trimmed = token.Trim();
            if (TryParseFormulaErrorLiteral(trimmed, out errorCode)) {
                return true;
            }

            Match functionMatch = SimpleFunctionFormulaRegex.Match(trimmed);
            if (functionMatch.Success) {
                string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                errorCode = function == "MATCH" || function == "XMATCH" || function == "XLOOKUP" || function == "VLOOKUP" || function == "HLOOKUP"
                    ? "#N/A"
                    : "#VALUE!";
                return true;
            }

            var binaryMatch = SimpleBinaryFormulaRegex.Match(trimmed);
            if (binaryMatch.Success
                && string.Equals(binaryMatch.Groups[2].Value, "/", StringComparison.Ordinal)
                && TryResolveNumericOperand(binaryMatch.Groups[3].Value, out double divisor)
                && Math.Abs(divisor) < double.Epsilon) {
                errorCode = "#DIV/0!";
                return true;
            }

            errorCode = string.Empty;
            return false;
        }

        private bool TryEvaluateFormulaOrNumeric(string token, out double result) {
            if (TryEvaluateFormula(token, out result) || TryResolveNumericOperand(token, out result)) {
                return true;
            }

            if (TryResolveFormulaArgument(token, out FormulaArgumentValue value) && value.Number.HasValue) {
                result = value.Number.Value;
                return true;
            }

            return false;
        }

        private bool TryEvaluateLogical(string args, bool useAnd, out bool result) {
            result = useAnd;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            foreach (string token in tokens) {
                if (!TryEvaluateCondition(token, out bool condition)) {
                    return false;
                }

                if (useAnd && !condition) {
                    result = false;
                    return true;
                }

                if (!useAnd && condition) {
                    result = true;
                    return true;
                }
            }

            result = useAnd;
            return true;
        }

        private bool TryEvaluateNot(string args, out bool result) {
            result = false;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1 || !TryEvaluateCondition(tokens[0], out bool condition)) {
                return false;
            }

            result = !condition;
            return true;
        }

        private bool TryEvaluateCountBlank(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1 || !TryResolveFormulaRange(tokens[0], out var values)) {
                return false;
            }

            result = values.Count(IsFormulaBlankValue);
            return true;
        }

        private bool TryEvaluateSubtotal(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || !TryGetWholeNumberArgument(tokens[0], out int functionCode)) {
                return false;
            }

            functionCode %= 100;
            if (functionCode != 1 && functionCode != 2 && functionCode != 3 && functionCode != 4 && functionCode != 5 && functionCode != 9) {
                return false;
            }

            var values = new List<FormulaArgumentValue>();
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            for (int index = 1; index < tokens.Count; index++) {
                if (!TryResolveFormulaRange(tokens[index], out var rangeValues, ref remainingCellBudget)) {
                    return false;
                }

                values.AddRange(rangeValues);
            }

            if (values.Count == 0) {
                return false;
            }

            if (functionCode == 3) {
                result = values.Count(value => !value.IsUnresolvedFormula && !IsFormulaBlankValue(value));
                return true;
            }

            var numbers = values.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (functionCode == 2) {
                result = numbers.Count;
                return true;
            }

            if (functionCode == 9) {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
                return false;
            }

            if (functionCode == 1) {
                result = numbers.Average();
            } else if (functionCode == 4) {
                result = numbers.Max();
            } else {
                result = numbers.Min();
            }

            return true;
        }

        private bool TryEvaluateConditionalAggregate(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            if (tokens.Count < 2 || tokens.Count > 3
                || !TryResolveFormulaRange(tokens[0], out var criteriaValues, ref remainingCellBudget)
                || !TryParseCriteria(tokens[1], out var criteria)) {
                return false;
            }

            var aggregateValues = criteriaValues;
            if (tokens.Count == 3) {
                if (!TryResolveFormulaRange(tokens[2], out aggregateValues, ref remainingCellBudget) || aggregateValues.Count != criteriaValues.Count) {
                    return false;
                }
            } else if (function == "AVERAGEIF") {
                aggregateValues = criteriaValues;
            }

            var matched = new List<FormulaArgumentValue>();
            for (int index = 0; index < criteriaValues.Count; index++) {
                if (MatchesCriteria(criteriaValues[index], criteria)) {
                    matched.Add(aggregateValues[index]);
                }
            }

            if (function == "COUNTIF") {
                result = matched.Count;
                return true;
            }

            var numbers = matched.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "SUMIF") {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
                return false;
            }

            result = numbers.Average();
            return true;
        }
    }
}
