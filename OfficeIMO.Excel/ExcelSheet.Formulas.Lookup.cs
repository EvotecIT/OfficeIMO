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
            var tokens = SplitFormulaArguments(args);
            if (function == "XLOOKUP") {
                return TryEvaluateXLookupValue(tokens, out result);
            }

            if (tokens.Count != 4
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryGetWholeNumberArgument(tokens[2], out int resultIndex)
                || !IsExactLookupMode(tokens[3])
                || !TryResolveFormulaRangeReference(tokens[1], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                || !TryGetFormulaRangeCellCount(r1, c1, r2, c2, out _)) {
                return false;
            }

            if (lookupValue.IsUnresolvedFormula) return false;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (function == "VLOOKUP") {
                int width = c2 - c1 + 1;
                if (resultIndex < 1 || resultIndex > width) {
                    return false;
                }

                int resultColumn = c1 + resultIndex - 1;
                for (int row = r1; row <= r2; row++) {
                    FormulaArgumentValue candidate = ResolveFormulaDependency(rangeSheet, row, c1);
                    if (candidate.IsUnresolvedFormula) return false;
                    if (candidate.IsError) { result = candidate; return true; }
                    if (!FormulaValuesEqual(candidate, lookupValue)) {
                        continue;
                    }

                    result = ResolveFormulaDependency(rangeSheet, row, resultColumn);
                    return result.HasValue;
                }

                result = FormulaArgumentValue.Error("#N/A");
                return true;
            }

            int height = r2 - r1 + 1;
            if (resultIndex < 1 || resultIndex > height) {
                return false;
            }

            int resultRow = r1 + resultIndex - 1;
            for (int column = c1; column <= c2; column++) {
                FormulaArgumentValue candidate = ResolveFormulaDependency(rangeSheet, r1, column);
                if (candidate.IsUnresolvedFormula) return false;
                if (candidate.IsError) { result = candidate; return true; }
                if (!FormulaValuesEqual(candidate, lookupValue)) {
                    continue;
                }

                result = ResolveFormulaDependency(rangeSheet, resultRow, column);
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
                || !TryResolveFormulaRange(tokens[2], out var returnValues, ref remainingCellBudget)
                || lookupValues.Count != returnValues.Count) {
                return false;
            }

            int matchMode = 0;
            if (tokens.Count >= 5 && !TryGetWholeNumberArgument(tokens[4], out matchMode)) {
                return false;
            }

            if (matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            int searchMode = 1;
            if (tokens.Count >= 6
                && (!TryGetWholeNumberArgument(tokens[5], out searchMode) || (searchMode != 1 && searchMode != -1))) {
                return false;
            }

            if (lookupValue.IsUnresolvedFormula) return false;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (TryFindLookupPosition(lookupValue, lookupValues, matchMode, searchMode, out int position, out FormulaArgumentValue searchFailure)) {
                var returnValue = returnValues[position - 1];
                result = returnValue;
                return result.HasValue;
            }

            if (searchFailure.IsUnresolvedFormula) return false;
            if (searchFailure.IsError) { result = searchFailure; return true; }
            if (tokens.Count >= 4) {
                if (!TryResolveFormulaArgument(tokens[3], out result) || result.IsUnresolvedFormula || !result.HasValue) return false;
                return true;
            }

            result = FormulaArgumentValue.Error("#N/A");
            return true;
        }

        private bool TryEvaluateSumProduct(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            var argumentSets = new List<List<double>>();
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            foreach (string token in tokens) {
                if (!TryResolveFormulaArgumentNumbers(token, out var values, ref remainingCellBudget) || values.Count == 0) {
                    return false;
                }

                if (argumentSets.Count > 0 && values.Count != argumentSets[0].Count) {
                    return false;
                }

                argumentSets.Add(values);
            }

            for (int index = 0; index < argumentSets[0].Count; index++) {
                double product = 1d;
                foreach (var values in argumentSets) {
                    product *= values[index];
                }

                result += product;
            }

            return true;
        }

    }
}
