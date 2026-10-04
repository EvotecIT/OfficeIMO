namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool TryEvaluateCombinations(double number, double chosen, out FormulaArgumentValue result) {
            result = default;
            if (number < 0 || chosen < 0 || number < chosen) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            double n = Math.Truncate(number);
            double k = Math.Min(Math.Truncate(chosen), n - Math.Truncate(chosen));
            // n >= 2*k, so C(n,k) >= 2^k. This also bounds evaluation work.
            if (k >= 1024) {
                result = FormulaArgumentValue.Error("#NUM!");
                return true;
            }
            var exact = System.Numerics.BigInteger.One;
            double combinations = 1;
            for (int i = 1; i <= (int)k; i++) {
                // Integer recurrence avoids factorial overflow and accumulated rounding.
                exact = exact * ((long)n - (i - 1)) / i;
                combinations = (double)exact;
                if (double.IsInfinity(combinations)) {
                    result = FormulaArgumentValue.Error("#NUM!");
                    return true;
                }
            }
            result = new FormulaArgumentValue(combinations, InvariantNumberText.Get(combinations));
            return true;
        }
    }
}
