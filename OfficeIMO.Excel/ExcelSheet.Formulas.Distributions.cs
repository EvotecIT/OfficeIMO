namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool IsDistributionFunction(string function) => function is
            "NORM.DIST" or "NORM.INV" or "NORM.S.DIST" or "NORM.S.INV"
            or "T.DIST" or "T.DIST.2T" or "T.DIST.RT" or "T.INV" or "T.INV.2T"
            or "CHISQ.DIST" or "CHISQ.DIST.RT" or "CHISQ.INV" or "CHISQ.INV.RT"
            or "NORMDIST" or "NORMINV" or "NORMSDIST" or "NORMSINV" or "TDIST" or "TINV" or "CHIDIST" or "CHIINV";

        private static int DistributionArgumentCount(string function) => function switch {
            "NORM.DIST" or "NORMDIST" => 4,
            "NORM.INV" or "NORMINV" or "T.DIST" or "CHISQ.DIST" or "TDIST" => 3,
            "NORM.S.INV" or "NORMSDIST" or "NORMSINV" => 1,
            _ => 2
        };

        // Numeric operands only: text, blanks and ranges stay unsupported rather than being coerced differently from Excel.
        private bool TryEvaluateDistributionValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            int count = DistributionArgumentCount(function);
            if (tokens.Count != count) return false;
            var numbers = new double[count];
            FormulaArgumentValue firstError = default;
            for (int index = 0; index < count; index++) {
                if (string.IsNullOrWhiteSpace(tokens[index])) return false;
                if (!TryResolveFormulaValueArgument(tokens[index], out FormulaArgumentValue value)) return false;
                if (value.IsError) {
                    if (!firstError.IsError) firstError = value;
                    continue;
                }
                if (!IsFiniteFormulaNumber(value)) return false;
                numbers[index] = value.Number!.Value;
            }
            if (firstError.IsError) {
                result = firstError;
                return true;
            }

            double outcome = EvaluateDistribution(function, numbers);
            result = double.IsNaN(outcome) || double.IsInfinity(outcome)
                ? FormulaArgumentValue.Error("#NUM!")
                : new FormulaArgumentValue(outcome, InvariantNumberText.Get(outcome));
            return true;
        }

        /// <summary>Returns NaN for every argument Excel rejects with #NUM!.</summary>
        private static double EvaluateDistribution(string function, double[] n) {
            switch (function) {
                case "NORM.DIST":
                case "NORMDIST": {
                    if (n[2] <= 0) return double.NaN;
                    double z = (n[0] - n[1]) / n[2];
                    return n[3] != 0 ? ExcelFormulaDistributions.NormalCumulative(z) : ExcelFormulaDistributions.NormalDensity(z) / n[2];
                }
                case "NORM.INV":
                case "NORMINV":
                    if (n[0] <= 0 || n[0] >= 1 || n[2] <= 0) return double.NaN;
                    return n[1] + n[2] * ExcelFormulaDistributions.NormalInverse(n[0]);
                case "NORM.S.DIST":
                    return n[1] != 0 ? ExcelFormulaDistributions.NormalCumulative(n[0]) : ExcelFormulaDistributions.NormalDensity(n[0]);
                case "NORMSDIST":
                    return ExcelFormulaDistributions.NormalCumulative(n[0]);
                case "NORM.S.INV":
                case "NORMSINV":
                    return n[0] <= 0 || n[0] >= 1 ? double.NaN : ExcelFormulaDistributions.NormalInverse(n[0]);
            }

            // The t and chi-squared functions truncate their degrees of freedom.
            double degrees = Math.Truncate(n[1]);
            switch (function) {
                case "T.DIST":
                    if (degrees < 1) return double.NaN;
                    return n[2] != 0
                        ? ExcelFormulaDistributions.StudentCumulative(n[0], degrees)
                        : ExcelFormulaDistributions.StudentDensity(n[0], degrees);
                case "T.DIST.2T":
                    if (degrees < 1 || n[0] < 0) return double.NaN;
                    return 2 * ExcelFormulaDistributions.StudentTail(n[0], degrees);
                case "T.DIST.RT":
                    if (degrees < 1) return double.NaN;
                    double tail = ExcelFormulaDistributions.StudentTail(n[0], degrees);
                    return n[0] > 0 ? tail : 1 - tail;
                case "TDIST": {
                    double tails = Math.Truncate(n[2]);
                    if (degrees < 1 || n[0] < 0 || (tails != 1 && tails != 2)) return double.NaN;
                    return tails * ExcelFormulaDistributions.StudentTail(n[0], degrees);
                }
                case "T.INV":
                    if (degrees < 1 || n[0] <= 0 || n[0] >= 1) return double.NaN;
                    return ExcelFormulaDistributions.StudentInverse(n[0], degrees);
                case "T.INV.2T":
                case "TINV":
                    if (degrees < 1 || n[0] <= 0 || n[0] > 1) return double.NaN;
                    return n[0] == 1 ? 0 : -ExcelFormulaDistributions.StudentInverse(n[0] / 2, degrees);
            }

            if (degrees < 1 || degrees > 1e10) return double.NaN;
            switch (function) {
                case "CHISQ.DIST":
                    if (n[0] < 0) return double.NaN;
                    return n[2] != 0
                        ? ExcelFormulaDistributions.ChiSquaredCumulative(n[0], degrees)
                        : ExcelFormulaDistributions.ChiSquaredDensity(n[0], degrees);
                case "CHISQ.DIST.RT":
                case "CHIDIST":
                    return n[0] < 0 ? double.NaN : ExcelFormulaDistributions.ChiSquaredUpper(n[0], degrees);
                case "CHISQ.INV":
                    return n[0] < 0 || n[0] > 1 ? double.NaN : ExcelFormulaDistributions.ChiSquaredInverse(n[0], degrees);
                case "CHISQ.INV.RT":
                case "CHIINV":
                    return n[0] < 0 || n[0] > 1 ? double.NaN : ExcelFormulaDistributions.ChiSquaredUpperInverse(n[0], degrees);
                default:
                    return double.NaN;
            }
        }
    }
}
