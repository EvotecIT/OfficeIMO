namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Scalar numeric subset: do not flatten a multi-cell range or discard a
        // nonnumeric operand to make an invalid signature appear valid.
        private bool TryEvaluateScalarMathValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            int count = function == "MOD" ? 2 : 1;
            if (tokens.Count != count || tokens.Any(string.IsNullOrWhiteSpace)) return false;
            var numbers = new double[count];
            FormulaArgumentValue firstError = default;
            for (int index = 0; index < count; index++) {
                if (!TryResolveFormulaArgument(tokens[index], out FormulaArgumentValue value) || value.IsUnresolvedFormula) return false;
                if (value.IsError) {
                    if (!firstError.IsError) firstError = value;
                    continue;
                }
                if (!IsFiniteFormulaNumber(value)) return false;
                numbers[index] = value.Number!.Value;
            }
            if (firstError.IsError) { result = firstError; return true; }
            if (function == "SQRT" && numbers[0] < 0) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            if (function == "MOD" && numbers[1] == 0) {
                result = FormulaArgumentValue.Error("#DIV/0!");
                return true;
            }
            double number = function == "INT" ? Math.Floor(numbers[0])
                : function == "SQRT" ? Math.Sqrt(numbers[0])
                : numbers[0] - numbers[1] * Math.Floor(numbers[0] / numbers[1]);
            result = double.IsNaN(number) || double.IsInfinity(number) ? FormulaArgumentValue.Error("#NUM!")
                : new FormulaArgumentValue(number, InvariantNumberText.Get(number));
            return true;
        }
    }
}
