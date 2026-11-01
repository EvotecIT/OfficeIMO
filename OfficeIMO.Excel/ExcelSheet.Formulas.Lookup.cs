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

            if (tokens.Count < 3 || tokens.Count > 4
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || lookupValue.IsUnresolvedFormula
                || !TryGetLookupPositionArgument(tokens[2], out int resultIndex, out FormulaArgumentValue indexError)) {
                return false;
            }
            FormulaArgumentValue modeValue = FormulaArgumentValue.Boolean(tokens.Count != 4 || tokens[3].Length > 0);
            if (tokens.Count == 4 && tokens[3].Length > 0
                && (!TryResolveFormulaArgument(tokens[3], out modeValue) || modeValue.IsUnresolvedFormula)) return false;
            if (!modeValue.IsError && !modeValue.Number.HasValue) return false;
            int budget = MaxResolvedFormulaRangeCells;
            if (!TryResolveLookupSource(tokens[1], ref budget, out FormulaLookupSource source)) return false;
            if (_formulaReferenceError.IsError) return true;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (indexError.IsError) { result = indexError; return true; }
            if (modeValue.IsError) { result = modeValue; return true; }
            bool vertical = function == "VLOOKUP", approximate = modeValue.Number != 0d;
            if (resultIndex < 1 || resultIndex > (vertical ? source.Columns : source.Rows)) {
                result = FormulaArgumentValue.Error(resultIndex < 1 ? "#VALUE!" : "#REF!");
                return true;
            }
            int matchMode = approximate ? -1 : !lookupValue.Number.HasValue ? 2 : 0;
            FormulaLookupSource vector = vertical ? source.Column(0) : source.Row(0);
            if (!TryFindLookupPosition(lookupValue, vector, matchMode, 1, out int position,
                out FormulaArgumentValue failure, sortedApproximate: approximate)) {
                if (failure.IsUnresolvedFormula) return false;
                result = failure.IsError ? failure : FormulaArgumentValue.Error("#N/A");
                return true;
            }
            result = NormalizeLookupResult(vertical ? source.At(position - 1, resultIndex - 1) : source.At(resultIndex - 1, position - 1));
            return !result.IsUnresolvedFormula;
        }

        private bool TryEvaluateXLookupValue(IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            if (!TrySelectXLookupSource(tokens, out FormulaLookupSource selected) || selected.Count != 1) return false;
            result = NormalizeLookupResult(selected[0]);
            return !result.IsUnresolvedFormula;
        }

        private bool TrySelectXLookupSource(IReadOnlyList<string> tokens, out FormulaLookupSource selected) {
            selected = null!;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            if (tokens.Count < 3 || tokens.Count > 6
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryResolveLookupSource(tokens[1], ref remainingCellBudget, out FormulaLookupSource lookupValues)
                || !TryResolveLookupSource(tokens[2], ref remainingCellBudget, out FormulaLookupSource returnValues)) {
                return false;
            }

            int matchMode = 0;
            FormulaArgumentValue matchError = default;
            if (tokens.Count >= 5 && tokens[4].Length > 0 && !TryGetLookupIntegerArgument(tokens[4], out matchMode, out matchError)) {
                return false;
            }

            if (!matchError.IsError && matchMode != -1 && matchMode != 0 && matchMode != 1 && matchMode != 2) {
                return false;
            }

            int searchMode = 1;
            FormulaArgumentValue modeError = default;
            if (tokens.Count >= 6 && tokens[5].Length > 0
                && (!TryGetLookupIntegerArgument(tokens[5], out searchMode, out modeError)
                    || (!modeError.IsError && searchMode != 1 && searchMode != -1 && searchMode != 2 && searchMode != -2))) {
                return false;
            }

            if (lookupValue.IsUnresolvedFormula) return false;
            FormulaArgumentValue error = _formulaReferenceError.IsError ? _formulaReferenceError
                : lookupValue.IsError ? lookupValue : matchError.IsError ? matchError : modeError.IsError ? modeError
                : lookupValues.Error.IsError ? lookupValues.Error : returnValues.Error;
            bool vertical = lookupValues.Columns == 1 && (lookupValues.Rows > 1 || returnValues.Rows == 1);
            if (!error.IsError && ((lookupValues.Rows != 1 && lookupValues.Columns != 1)
                || (vertical ? lookupValues.Rows != returnValues.Rows : lookupValues.Columns != returnValues.Columns)))
                error = FormulaArgumentValue.Error("#VALUE!");
            if (error.IsError) { selected = LookupScalarSource(error); return true; }
            if (TryFindLookupPosition(lookupValue, lookupValues, matchMode, searchMode, out int position, out FormulaArgumentValue searchFailure)) {
                selected = vertical ? returnValues.Row(position - 1) : returnValues.Column(position - 1);
                return true;
            }

            if (searchFailure.IsUnresolvedFormula) return false;
            if (searchFailure.IsError) { selected = LookupScalarSource(searchFailure); return true; }
            if (tokens.Count >= 4 && tokens[3].Length > 0) {
                if (!TryResolveFormulaArgument(tokens[3], out FormulaArgumentValue fallback) || fallback.IsUnresolvedFormula) return false;
                selected = LookupScalarSource(NormalizeLookupResult(fallback));
                return true;
            }

            selected = LookupScalarSource(FormulaArgumentValue.Error("#N/A"));
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

        // Row/column positions coerce numeric text and truncate fractions. Modes
        // retain their separate discrete-value qualification above.
        private bool TryGetLookupPositionArgument(string token, out int number, out FormulaArgumentValue error) {
            number = 0;
            error = default;
            if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value) || value.IsUnresolvedFormula) return false;
            if (value.IsError) { error = value; return true; }
            double position = value.Number ?? 0d;
            if (value.HasValue && !value.Number.HasValue
                && !double.TryParse(value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out position)) {
                error = FormulaArgumentValue.Error("#VALUE!");
                return true;
            }
            if (!IsFinite(position)) { error = FormulaArgumentValue.Error("#VALUE!"); return true; }
            number = position < 0 ? -1 : position >= int.MaxValue ? int.MaxValue : (int)Math.Truncate(position);
            return true;
        }

        // Excel returns zero for a selected blank cell, while an unresolved formula
        // must keep its state so an error handler cannot invent a fallback.
        private static FormulaArgumentValue NormalizeLookupResult(FormulaArgumentValue value) =>
            !value.HasValue && !value.IsUnresolvedFormula ? new FormulaArgumentValue(0d, "0") : value;

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
