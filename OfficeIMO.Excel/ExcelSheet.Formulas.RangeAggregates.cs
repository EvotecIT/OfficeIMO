namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool IsRangeAggregateFunction(string function) =>
            function is "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT";

        /// <summary>
        /// Folds an aggregate over one plain range without materializing its cells. Any other argument shape
        /// returns <paramref name="handled"/> false and takes the general argument route, which sets the contract.
        /// </summary>
        private bool TryEvaluateRangeAggregate(string function, string args, out double result, out FormulaArgumentValue error, out bool handled) {
            result = 0;
            error = default;
            handled = false;
            if (args.IndexOf('!') >= 0
                || !ExcelA1.TryParseRange(args.Replace("$", string.Empty), out int r1, out int c1, out int r2, out int c2)
                || !TryGetFormulaRangeCellCount(r1, c1, r2, c2, out _)) {
                return false;
            }

            handled = true;
            bool unresolved = false;
            FormulaArgumentValue firstError = default;
            int count = 0;
            double sum = 0, minimum = double.PositiveInfinity, maximum = double.NaN;
            bool minimumIsNaN = false;
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    // Every dependency resolves even after a failure: later cells can still raise the guard.
                    FormulaArgumentValue value = ResolveFormulaDependency(this, row, column);
                    if (value.IsUnresolvedFormula) unresolved = true;
                    if (value.IsError && !firstError.IsError) firstError = value;
                    if (!value.IsNumericAggregateValue) continue;
                    double number = value.Number!.Value;
                    count++;
                    sum += number;
                    if (double.IsNaN(number)) minimumIsNaN = true;
                    else {
                        if (number < minimum) minimum = number;
                        if (double.IsNaN(maximum) || number > maximum) maximum = number;
                    }
                }
            }

            if (unresolved) return false;
            if (_formulaReferenceError.IsError) { error = _formulaReferenceError; return false; }
            if (function == "COUNT") { result = count; return true; }
            if (firstError.IsError) { error = firstError; return false; }
            if (count == 0) {
                if (function == "AVERAGE") { error = FormulaArgumentValue.Error("#DIV/0!"); return false; }
                return true;
            }

            result = function switch {
                "SUM" => sum,
                "AVERAGE" => sum / count,
                "MIN" => minimumIsNaN ? double.NaN : minimum,
                _ => maximum
            };
            return true;
        }
    }
}
