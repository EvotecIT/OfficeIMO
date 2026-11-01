using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateLookupFunction(string function, string args, out double result) {
            result = 0;
            if (!TryEvaluateLookupValue(function, args, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateLookupValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (function == "XLOOKUP") {
                return TryEvaluateXLookupValue(tokens, out result);
            }

            if (tokens.Count != 4
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || lookupValue.IsUnresolvedFormula
                || !TryGetLookupIntegerArgument(tokens[2], out int resultIndex, out FormulaArgumentValue indexError)
                || !TryResolveFormulaArgument(tokens[3], out FormulaArgumentValue modeValue)
                || modeValue.IsUnresolvedFormula
                || (!modeValue.IsError && modeValue.Number != 0d)) {
                return false;
            }
            if (!TryResolveFormulaRangeReference(tokens[1], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2))
                return _formulaReferenceError.IsError;
            if (!TryGetFormulaRangeCellCount(r1, c1, r2, c2, out _)) return false;
            if (!IsSupportedExactLookupValue(lookupValue)) return false;

            if (lookupValue.IsUnresolvedFormula) return false;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (indexError.IsError) { result = indexError; return true; }
            if (modeValue.IsError) { result = modeValue; return true; }
            if (function == "VLOOKUP") {
                int width = c2 - c1 + 1;
                if (resultIndex < 1 || resultIndex > width) {
                    result = FormulaArgumentValue.Error(resultIndex < 1 ? "#VALUE!" : "#REF!");
                    return true;
                }

                int resultColumn = c1 + resultIndex - 1;
                for (int row = r1; row <= r2; row++) {
                    FormulaArgumentValue candidate = ResolveFormulaDependency(rangeSheet, row, c1);
                    if (candidate.IsUnresolvedFormula) return false;
                    if (candidate.IsError) { result = candidate; return true; }
                    if (!FormulaValuesEqual(candidate, lookupValue)) {
                        continue;
                    }

                    result = NormalizeLookupResult(ResolveFormulaDependency(rangeSheet, row, resultColumn));
                    return result.HasValue;
                }

                result = FormulaArgumentValue.Error("#N/A");
                return true;
            }

            int height = r2 - r1 + 1;
            if (resultIndex < 1 || resultIndex > height) {
                result = FormulaArgumentValue.Error(resultIndex < 1 ? "#VALUE!" : "#REF!");
                return true;
            }

            int resultRow = r1 + resultIndex - 1;
            for (int column = c1; column <= c2; column++) {
                FormulaArgumentValue candidate = ResolveFormulaDependency(rangeSheet, r1, column);
                if (candidate.IsUnresolvedFormula) return false;
                if (candidate.IsError) { result = candidate; return true; }
                if (!FormulaValuesEqual(candidate, lookupValue)) {
                    continue;
                }

                result = NormalizeLookupResult(ResolveFormulaDependency(rangeSheet, resultRow, column));
                return result.HasValue;
            }

            result = FormulaArgumentValue.Error("#N/A");
            return true;
        }

        private bool TryEvaluateXLookupValue(IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            if (tokens.Count < 3 || tokens.Count > 6
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryResolveFormulaRange(tokens[1], out var lookupValues, ref remainingCellBudget)
                || !TryResolveFormulaRange(tokens[2], out var returnValues, ref remainingCellBudget)) {
                return false;
            }

            int matchMode = 0;
            FormulaArgumentValue matchError = default;
            if (tokens.Count >= 5 && tokens[4].Length > 0 && !TryGetLookupIntegerArgument(tokens[4], out matchMode, out matchError)) {
                return false;
            }

            if (!matchError.IsError && matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            int searchMode = 1;
            FormulaArgumentValue modeError = default;
            if (tokens.Count >= 6 && tokens[5].Length > 0
                && (!TryGetLookupIntegerArgument(tokens[5], out searchMode, out modeError)
                    || (!modeError.IsError && searchMode != 1 && searchMode != -1))) {
                return false;
            }

            if (lookupValue.IsUnresolvedFormula) return false;
            if (_formulaReferenceError.IsError) return true;
            if (lookupValues.Count != returnValues.Count) return false;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (matchError.IsError) { result = matchError; return true; }
            if (modeError.IsError) { result = modeError; return true; }
            if (TryFindLookupPosition(lookupValue, lookupValues, matchMode, searchMode, out int position, out FormulaArgumentValue searchFailure)) {
                result = NormalizeLookupResult(returnValues[position - 1]);
                return result.HasValue;
            }

            if (searchFailure.IsUnresolvedFormula) return false;
            if (searchFailure.IsError) { result = searchFailure; return true; }
            if (tokens.Count >= 4 && tokens[3].Length > 0) {
                if (!TryResolveFormulaArgument(tokens[3], out result) || result.IsUnresolvedFormula || !result.HasValue) return false;
                return true;
            }

            result = FormulaArgumentValue.Error("#N/A");
            return true;
        }

        /// <summary>
        /// Resolves an integral lookup operand without discarding a typed Excel error.
        /// Unsupported or unresolved operands still leave the owning formula unevaluated.
        /// </summary>
        private bool TryGetLookupIntegerArgument(string token, out int number, out FormulaArgumentValue error) {
            number = 0;
            error = default;
            if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value) || value.IsUnresolvedFormula) return false;
            if (value.IsError) { error = value; return true; }
            return value.Number.HasValue && TryGetWholeNumber(value.Number.Value, out number);
        }

        // Excel returns zero for a selected blank cell, while an unresolved formula
        // must keep its state so an error handler cannot invent a fallback.
        private static FormulaArgumentValue NormalizeLookupResult(FormulaArgumentValue value) =>
            !value.HasValue && !value.IsUnresolvedFormula ? new FormulaArgumentValue(0d, "0") : value;

        // Legacy exact lookups interpret these characters as wildcard syntax.
        // Until that mode is qualified, never cache a literal-string search result.
        private static bool IsSupportedExactLookupValue(FormulaArgumentValue value) =>
            value.IsError || value.Number.HasValue || value.Text == null
                || (value.Text.IndexOf('*') < 0 && value.Text.IndexOf('?') < 0 && value.Text.IndexOf('~') < 0);

        private bool TryEvaluateSumProduct(string args, out double result) {
            result = 0;
            if (!TryEvaluateSumProductValue(args, out FormulaArgumentValue value) || !value.Number.HasValue) return false;
            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateSumProductValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            var argumentSets = new List<List<FormulaArgumentValue>>();
            FormulaArgumentValue firstError = default;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            foreach (string token in tokens) {
                if (!TryResolveFormulaRange(token, out var values, ref remainingCellBudget, out bool isReference)) {
                    if (isReference || !TryResolveFormulaArgument(token, out FormulaArgumentValue scalar)) return false;
                    values = new List<FormulaArgumentValue> { scalar };
                }
                foreach (FormulaArgumentValue value in values) {
                    if (value.IsUnresolvedFormula) return false;
                    if (value.IsError) { if (!firstError.IsError) firstError = value; }
                    else if (!value.IsNumericAggregateValue || !IsFinite(value.Number!.Value)) return false;
                }
                argumentSets.Add(values);
            }
            if (_formulaReferenceError.IsError) return true;
            if (argumentSets.Any(values => values.Count == 0 || values.Count != argumentSets[0].Count)) return false;
            if (firstError.IsError) { result = firstError; return true; }

            double sum = 0;
            for (int index = 0; index < argumentSets[0].Count; index++) {
                double product = 1d;
                foreach (var values in argumentSets) {
                    product *= values[index].Number!.Value;
                }

                sum += product;
            }
            result = IsFinite(sum) ? new FormulaArgumentValue(sum, InvariantNumberText.Get(sum)) : FormulaArgumentValue.Error("#NUM!");
            return true;
        }

    }
}
