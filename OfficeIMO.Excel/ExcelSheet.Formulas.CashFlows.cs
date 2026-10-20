namespace OfficeIMO.Excel {

    public partial class ExcelSheet {
        /// <summary>Evaluates periodic and dated cash flows through the shared typed argument resolver.</summary>
        private bool TryEvaluateCashFlowValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (function == "RATE") return TryEvaluateRateValue(tokens, out result);
            bool dated = function != "IRR", discount = function == "XNPV";
            int required = dated ? (discount ? 3 : 2) : 1;
            if (tokens.Count < required || tokens.Count > required + (discount ? 0 : 1)) return false;
            int budget = MaxResolvedFormulaRangeCells;
            int valuesIndex = discount ? 1 : 0;
            if (!TryResolveFinancialSeries(tokens[valuesIndex], ref budget, out var source)) return false;
            if (source.Any(value => value.IsUnresolvedFormula)) return false;
            FormulaArgumentValue error = source.FirstOrDefault(value => value.IsError);
            if (error.IsError) {
                result = !dated && !TryParseFormulaErrorLiteral(tokens[valuesIndex].Trim(), out _)
                    ? FormulaArgumentValue.Error("#VALUE!") : error;
                return true;
            }
            var values = new List<double>();
            foreach (var value in source) {
                if (!dated && (!value.IsNumericAggregateValue || value.IsBoolean)) continue;
                if (dated && value.HasValue && (!value.Number.HasValue || value.IsBoolean)) {
                    result = FormulaArgumentValue.Error("#VALUE!"); return true;
                }
                if (discount && !value.HasValue) continue;
                values.Add(value.Number ?? 0d);
            }
            double[] periods = new double[values.Count];
            if (dated) {
                if (!TryResolveFinancialSeries(tokens[valuesIndex + 1], ref budget, out var dates)) return false;
                if (dates.Any(value => value.IsUnresolvedFormula)) return false;
                error = dates.FirstOrDefault(value => value.IsError);
                if (error.IsError) { result = error; return true; }
                if (dates.Count != values.Count) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
                double firstDate = 0;
                for (int index = 0; index < dates.Count; index++) {
                    var date = dates[index];
                    if (!date.Number.HasValue || date.IsBoolean || !IsFinite(date.Number.Value)) {
                        result = FormulaArgumentValue.Error("#VALUE!"); return true;
                    }
                    if (date.Number.Value < 0 || !TryGetDateFromSerial(date.Number.Value, out _)) {
                        result = FormulaArgumentValue.Error("#NUM!"); return true;
                    }
                    double serial = Math.Truncate(date.Number.Value);
                    if (index == 0) firstDate = serial;
                    if (serial < firstDate) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
                    periods[index] = (serial - firstDate) / 365d;
                }
            } else for (int index = 0; index < periods.Length; index++) periods[index] = index;
            double rate = discount ? 0 : 0.1;
            int scalarIndex = discount ? 0 : required;
            if (discount && string.IsNullOrWhiteSpace(tokens[0])) {
                result = FormulaArgumentValue.Error("#N/A"); return true;
            }
            if (scalarIndex < tokens.Count && !string.IsNullOrWhiteSpace(tokens[scalarIndex])) {
                if (!TryFinancialScalar(tokens[scalarIndex], out rate, out error)) return false;
                if (error.IsError) { result = error; return true; }
            }
            if (values.Count == 0 || rate <= -1 || discount && rate <= 0 || values.Any(value => !IsFinite(value))) {
                result = FormulaArgumentValue.Error(!dated && rate <= -1 ? "#VALUE!" : "#NUM!"); return true;
            }
            if (discount) {
                CashFlowBalance(values, periods, rate, out double balance, out _, out _);
                result = FinancialNumberOrError(balance); return true;
            }
            if (!values.Any(value => value > 0) || !values.Any(value => value < 0)) {
                result = FormulaArgumentValue.Error("#NUM!"); return true;
            }
            double scale = values.Max(value => Math.Abs(value));
            double[] normalized = values.Select(value => value / scale).ToArray();
            bool solved = TrySolveFinancialRate(rate, dated ? 100 : 20, dated ? 1e-10 : 1e-7,
                (double candidate, out double balance, out double derivative, out double magnitude) =>
                    CashFlowBalance(normalized, periods, candidate, out balance, out derivative, out magnitude), out double solution);
            result = solved ? FinancialNumberOrError(solution) : FormulaArgumentValue.Error("#NUM!");
            return true;
        }

        private bool TryEvaluateRateValue(IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            if (tokens.Count < 3 || tokens.Count > 6) return false;
            var numbers = new double[] { 0, 0, 0, 0, 0, 0.1 };
            FormulaArgumentValue firstError = default;
            for (int index = 0; index < tokens.Count; index++) {
                if (string.IsNullOrWhiteSpace(tokens[index])) {
                    if (index == 1 && tokens.Count >= 4 || index >= 3) continue;
                    return false;
                }
                if (!TryFinancialScalar(tokens[index], out numbers[index], out var error)) return false;
                if (error.IsError && !firstError.IsError) firstError = error;
            }
            if (firstError.IsError) { result = firstError; return true; }
            double n = numbers[0], payment = numbers[1], present = numbers[2], future = numbers[3], type = numbers[4] == 0 ? 0 : 1;
            if (n <= 0 || numbers[5] <= -1) {
                result = FormulaArgumentValue.Error("#NUM!"); return true;
            }
            double scale = Math.Max(Math.Abs(present), Math.Max(Math.Abs(payment), Math.Abs(future)));
            if (scale == 0) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
            present /= scale; payment /= scale; future /= scale;
            bool solved = TrySolveFinancialRate(numbers[5], 20, 1e-7,
                (double rate, out double balance, out double derivative, out double magnitude) => {
                    if (!TryFinancialAnnuityFactors(rate, n, out double factor, out double annuity, out double annuityDerivative)) {
                        balance = derivative = magnitude = double.NaN; return;
                    }
                    double factorDerivative = n * factor / (1 + rate);
                    balance = present * factor + payment * (1 + type * rate) * annuity + future;
                    derivative = present * factorDerivative + payment * (type * annuity + (1 + type * rate) * annuityDerivative);
                    magnitude = Math.Abs(present * factor) + Math.Abs(payment * (1 + type * rate) * annuity) + Math.Abs(future);
                }, out double solution);
            result = solved ? FinancialNumberOrError(solution) : FormulaArgumentValue.Error("#NUM!");
            return true;
        }

        private delegate void FinancialBalance(double rate, out double balance, out double derivative, out double magnitude);

        private static bool TrySolveFinancialRate(double guess, int iterations, double tolerance,
            FinancialBalance evaluate, out double result) {
            if (TrySolveFinancialRateAttempt(guess, iterations, tolerance, evaluate, out result)) return true;
            // A poor guess can approach the zero-discount limit at infinity instead of a finite root.
            // Retry once from the documented default; each attempt keeps the same finite iteration budget.
            return guess != 0.1 && TrySolveFinancialRateAttempt(0.1, iterations, tolerance, evaluate, out result);
        }

        private static bool TrySolveFinancialRateAttempt(double guess, int iterations, double tolerance,
            FinancialBalance evaluate, out double result) {
            result = guess;
            for (int iteration = 0; iteration < iterations; iteration++) {
                evaluate(result, out double balance, out double derivative, out double magnitude);
                if (!IsFinite(balance) || !IsFinite(derivative) || !IsFinite(magnitude) || magnitude <= 0) return false;
                if (derivative == 0) return balance == 0;
                double step = balance / derivative;
                double next = result - step;
                if (!IsFinite(next)) return false;
                // Keep a Newton step inside the real discounting domain without selecting a remote root.
                if (next <= -1) next = (result - 1) / 2;
                // Check the computed correction itself: subtracting it from a large rate can round it away.
                double change = Math.Abs(step);
                result = next;
                if (change <= tolerance) {
                    evaluate(result, out balance, out derivative, out magnitude);
                    return IsFinite(balance) && IsFinite(derivative) && IsFinite(magnitude) && magnitude > 0
                        && Math.Abs(balance) / magnitude <= 1e-8;
                }
            }
            return false;
        }

        private static void CashFlowBalance(IReadOnlyList<double> values, IReadOnlyList<double> periods,
            double rate, out double balance, out double derivative, out double magnitude) {
            balance = derivative = magnitude = 0;
            double balanceCorrection = 0, derivativeCorrection = 0;
            for (int index = 0; index < values.Count; index++) {
                double discounted = values[index] / Math.Pow(1 + rate, periods[index]);
                magnitude += Math.Abs(discounted);
                AddFinancialTerm(ref balance, ref balanceCorrection, discounted);
                AddFinancialTerm(ref derivative, ref derivativeCorrection, -periods[index] * discounted / (1 + rate));
            }
        }

        private static void AddFinancialTerm(ref double sum, ref double correction, double term) {
            double adjusted = term - correction, next = sum + adjusted;
            correction = (next - sum) - adjusted;
            sum = next;
        }

        private static FormulaArgumentValue FinancialNumberOrError(double value) => IsFinite(value)
            ? new FormulaArgumentValue(value, InvariantNumberText.Get(value)) : FormulaArgumentValue.Error("#NUM!");
    }
}
