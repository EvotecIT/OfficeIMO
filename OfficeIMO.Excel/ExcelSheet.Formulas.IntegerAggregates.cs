namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Evaluates integer aggregates within the consecutive binary64 integer domain.
        /// Unqualified operands retain caches; supported domain failures are typed errors.
        /// </summary>
        private bool TryEvaluateIntegerAggregateValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 1 || tokens.Count > 255 || tokens.Any(string.IsNullOrWhiteSpace)
                || !TryResolveFormulaArguments(args, out List<FormulaArgumentValue> values,
                    skipBlankScalarReferences: true)) return false;
            if (values.Count == 0) {
                result = FormulaArgumentValue.Error("#VALUE!");
                return true;
            }

            const ulong maximumExactInteger = 9007199254740991UL;
            var operands = new List<ulong>(values.Count);
            FormulaArgumentValue firstError = default;
            foreach (FormulaArgumentValue value in values) {
                if (value.IsUnresolvedFormula) return false;
                if (value.IsError) {
                    if (!firstError.IsError) firstError = value;
                    continue;
                }
                double number = 0;
                if (value.IsBoolean) {
                    if (!firstError.IsError) firstError = FormulaArgumentValue.Error("#VALUE!");
                    continue;
                }
                if (value.Number.HasValue) number = value.Number.Value;
                else if (value.Text != null && !double.TryParse(value.Text,
                    System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number)) {
                    if (!firstError.IsError) firstError = FormulaArgumentValue.Error("#VALUE!");
                    continue;
                }
                // Validate every operand before exposing an error to IFERROR. Blank
                // range positions are zero; direct blanks were filtered above.
                // Finite negative operands are #NUM!.
                if (double.IsNaN(number) || double.IsInfinity(number)
                    || number > maximumExactInteger || number < -(double)maximumExactInteger) return false;
                if (number < 0) {
                    if (!firstError.IsError) firstError = FormulaArgumentValue.Error("#NUM!");
                    continue;
                }
                operands.Add((ulong)Math.Truncate(number));
            }
            if (firstError.IsError) { result = firstError; return true; }

            ulong aggregate = function == "GCD" ? 0UL : 1UL;
            bool overflow = false, hasZero = false;
            foreach (ulong operand in operands) {
                if (function == "GCD") aggregate = GreatestCommonDivisor(aggregate, operand);
                else if (operand == 0) hasZero = true;
                else if (!overflow) {
                    ulong factor = operand / GreatestCommonDivisor(aggregate, operand);
                    if (factor > maximumExactInteger / aggregate) overflow = true;
                    else aggregate *= factor;
                }
            }
            // A zero anywhere in LCM wins over an intermediate product overflow.
            // Keep scanning validation inputs rather than making argument order matter.
            if (function == "LCM" && hasZero) aggregate = 0;
            else if (overflow) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
            result = new FormulaArgumentValue(aggregate, InvariantNumberText.Get(aggregate));
            return true;
        }

        private static ulong GreatestCommonDivisor(ulong left, ulong right) {
            while (right != 0) {
                ulong remainder = left % right;
                left = right;
                right = remainder;
            }
            return left;
        }
    }
}
