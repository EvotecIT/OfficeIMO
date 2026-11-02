namespace OfficeIMO.Excel {

    public partial class ExcelSheet {
        /// <summary>Splits a constant-rate annuity payment using the existing PMT/FV financial owner.</summary>
        private bool TryEvaluatePaymentSplitValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 4 || tokens.Count > 6) return false;
            var numbers = new double[6];
            FormulaArgumentValue firstError = default;
            for (int index = 0; index < tokens.Count; index++) {
                if (string.IsNullOrWhiteSpace(tokens[index])) {
                    if (index < 4) return false;
                    continue;
                }
                if (!TryResolveFormulaValueArgument(tokens[index], out FormulaArgumentValue value)) return false;
                if (value.IsError) { if (!firstError.IsError) firstError = value; continue; }
                if (!IsFiniteFormulaNumber(value)) return false;
                numbers[index] = value.Number!.Value;
            }
            if (firstError.IsError) { result = firstError; return true; }
            double rate = numbers[0], period = numbers[1], periods = numbers[2], present = numbers[3], future = numbers[4], type = numbers[5] == 0 ? 0 : 1;
            if (period < 1 || period > periods || periods <= 0 || rate <= -1) {
                result = FormulaArgumentValue.Error("#NUM!"); return true;
            }
            if (!TryEvaluatePmt(new[] { rate, periods, present, future, type }, out double payment)
                || !TryEvaluateFv(new[] { rate, period - 1, payment, present, type }, out double balance)) {
                result = FormulaArgumentValue.Error("#NUM!"); return true;
            }
            double interest = type == 1 && period == 1 ? 0 : balance * rate / (1 + rate * type);
            double valueResult = function == "IPMT" ? interest : payment - interest;
            result = IsFinite(valueResult) ? new FormulaArgumentValue(valueResult, InvariantNumberText.Get(valueResult)) : FormulaArgumentValue.Error("#NUM!");
            return true;
        }
    }
}
