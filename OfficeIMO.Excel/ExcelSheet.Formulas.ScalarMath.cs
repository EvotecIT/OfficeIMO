namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Scalar numeric subset: do not flatten a multi-cell range or discard a
        // nonnumeric operand to make an invalid signature appear valid.
        private bool TryEvaluateScalarMathValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            bool directional = function is "TRUNC" or "ROUNDUP" or "ROUNDDOWN";
            bool multiple = function is "CEILING" or "FLOOR";
            int minimum = multiple || function is "MOD" or "ROUNDUP" or "ROUNDDOWN" ? 2 : 1;
            int maximum = function is "TRUNC" or "LOG" ? 2 : minimum;
            int count = tokens.Count;
            if (count < minimum || count > maximum || (!multiple && tokens.Any(string.IsNullOrWhiteSpace))) return false;
            var numbers = new double[count];
            FormulaArgumentValue firstError = default;
            bool digitsAreError = false;
            for (int index = 0; index < count; index++) {
                // Explicitly omitted CEILING/FLOOR operands have the numeric zero value.
                if (multiple && string.IsNullOrWhiteSpace(tokens[index])) continue;
                if (!TryResolveFormulaArgument(tokens[index], out FormulaArgumentValue value) || value.IsUnresolvedFormula) return false;
                if (value.IsError) {
                    if (!firstError.IsError) firstError = value;
                    if (index == 1) digitsAreError = true;
                    continue;
                }
                if (!IsFiniteFormulaNumber(value)) return false;
                numbers[index] = value.Number!.Value;
            }
            int digits = 0;
            if (directional && count == 2 && !digitsAreError && !TryGetSupportedDecimalPlaces(numbers[1], out digits)) return false;
            if (firstError.IsError) { result = firstError; return true; }
            if (function is "EVEN" or "ODD") {
                // Round the magnitude before checking parity: dividing tiny values by
                // two first can underflow to zero. Large binary64 integers are even.
                double magnitude = Math.Ceiling(Math.Abs(numbers[0]));
                bool odd = function == "ODD";
                // No odd integer beyond this value is exactly representable in binary64.
                if (odd && magnitude > 9007199254740991d) return false;
                if ((magnitude % 2 == 0) == odd) magnitude += 1;
                double rounded = numbers[0] < 0 ? -magnitude : magnitude;
                result = new FormulaArgumentValue(rounded, InvariantNumberText.Get(rounded));
                return true;
            }
            if (function == "SQRT" && numbers[0] < 0) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            if (function == "MOD" && numbers[1] == 0) {
                result = FormulaArgumentValue.Error("#DIV/0!");
                return true;
            }
            if (multiple) {
                if (numbers[0] == 0 || (function == "CEILING" && numbers[1] == 0)) {
                    result = new FormulaArgumentValue(0, "0");
                    return true;
                }
                if (numbers[1] == 0 || (numbers[0] > 0 && numbers[1] < 0)) {
                    result = FormulaArgumentValue.Error(numbers[1] == 0 ? "#DIV/0!" : "#NUM!");
                    return true;
                }
            }
            double logarithmBase = count == 2 ? numbers[1] : 10;
            if ((function is "LN" or "LOG" or "LOG10") &&
                (numbers[0] <= 0 || (function == "LOG" && logarithmBase <= 0))) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            if (function == "LOG" && logarithmBase == 1) {
                result = FormulaArgumentValue.Error("#DIV/0!");
                return true;
            }
            double number = multiple ? RoundAtMultiple(numbers[0], numbers[1], ceiling: function == "CEILING")
                : directional ? RoundDirectionalAtDigits(numbers[0], digits, awayFromZero: function == "ROUNDUP")
                : function == "EXP" ? Math.Exp(numbers[0])
                : function == "LN" ? Math.Log(numbers[0])
                : function == "LOG10" || (function == "LOG" && logarithmBase == 10) ? Math.Log10(numbers[0])
                : function == "LOG" ? Math.Log(numbers[0], logarithmBase)
                : function == "SIGN" ? Math.Sign(numbers[0])
                : function == "INT" ? Math.Floor(numbers[0])
                : function == "SQRT" ? Math.Sqrt(numbers[0])
                : numbers[0] - numbers[1] * Math.Floor(numbers[0] / numbers[1]);
            result = double.IsNaN(number) || double.IsInfinity(number) ? FormulaArgumentValue.Error("#NUM!")
                : new FormulaArgumentValue(number, InvariantNumberText.Get(number));
            return true;
        }
    }
}
