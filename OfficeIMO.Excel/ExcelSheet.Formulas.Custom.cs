using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateCustomFormulaFunction(string formula, out FormulaArgumentValue result) {
            result = default;
            Match match = AnyFunctionFormulaRegex.Match(formula);
            if (!match.Success) {
                return false;
            }

            string functionName = match.Groups[1].Value.ToUpperInvariant();
            if (!_excelDocument.Calculation.TryGetCustomFunction(functionName, out ExcelCustomFormulaFunction? function)
                || function == null
                || !TryResolveFormulaArguments(match.Groups[2].Value, out List<FormulaArgumentValue> arguments)
                || arguments.Any(argument => argument.IsUnresolvedFormula)) {
                return false;
            }

            var customArguments = new ExcelFormulaValue[arguments.Count];
            for (int index = 0; index < arguments.Count; index++) {
                customArguments[index] = ToCustomFormulaValue(arguments[index]);
            }

            var context = new ExcelCustomFormulaFunctionContext(
                _excelDocument,
                this,
                functionName,
                _formulaEvaluationCellReference);
            ExcelFormulaValue? customResult = function(context, customArguments);
            return customResult.HasValue && TryConvertCustomFormulaValue(customResult.Value, out result);
        }

        private static ExcelFormulaValue ToCustomFormulaValue(FormulaArgumentValue value) {
            if (value.IsError) {
                return ExcelFormulaValue.FromError(value.ErrorCode ?? "#VALUE!");
            }

            if (value.Number.HasValue) {
                return ExcelFormulaValue.FromNumber(value.Number.Value);
            }

            return value.Text == null
                ? ExcelFormulaValue.Blank
                : ExcelFormulaValue.FromText(value.Text);
        }

        private static bool TryConvertCustomFormulaValue(ExcelFormulaValue value, out FormulaArgumentValue result) {
            switch (value.Kind) {
                case ExcelFormulaValueKind.Blank:
                    result = new FormulaArgumentValue(null, string.Empty);
                    return true;
                case ExcelFormulaValueKind.Number:
                    if (double.IsNaN(value.Number) || double.IsInfinity(value.Number)) {
                        result = default;
                        return false;
                    }

                    result = new FormulaArgumentValue(value.Number, InvariantNumberText.Get(value.Number));
                    return true;
                case ExcelFormulaValueKind.Text:
                    result = new FormulaArgumentValue(null, value.Text ?? string.Empty);
                    return true;
                case ExcelFormulaValueKind.Error:
                    if (string.IsNullOrWhiteSpace(value.Text)) {
                        result = default;
                        return false;
                    }

                    result = FormulaArgumentValue.Error(value.Text!);
                    return true;
                default:
                    result = default;
                    return false;
            }
        }
    }
}
