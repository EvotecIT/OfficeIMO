namespace OfficeIMO.Excel {

    public partial class ExcelSheet {
        /// <summary>Shares stable compound-interest and annuity factors across payment and inverse-rate calculations.</summary>
        private static bool TryFinancialAnnuityFactors(double rate, double periods,
            out double factor, out double annuity, out double annuityDerivative) {
            factor = rate > -1 ? Math.Exp(periods * FinancialLog1p(rate)) : Math.Pow(1 + rate, periods);
            annuity = annuityDerivative = 0;
            if (Math.Abs(rate) < 0.0001 && Math.Abs(rate * periods) < 0.0001) {
                // Generalized binomial series preserves the zero-rate limit, including fractional periods.
                annuity = periods;
                double coefficient = periods, power = 1, previousPower = 1;
                for (int order = 2; order <= 10; order++) {
                    coefficient *= (periods - order + 1) / order;
                    power *= rate;
                    annuity += coefficient * power;
                    annuityDerivative += (order - 1) * coefficient * previousPower;
                    previousPower = power;
                }
            } else if (rate != 0) {
                double difference = rate > -1 ? FinancialExpm1(periods * FinancialLog1p(rate)) : factor - 1;
                annuity = difference / rate;
                annuityDerivative = (periods * factor / (1 + rate) - annuity) / rate;
            }
            return IsFinite(factor) && IsFinite(annuity);
        }

        private static double FinancialLog1p(double value) {
            if (Math.Abs(value) >= 0.0001) return Math.Log(1 + value);
            double sum = value, power = value;
            for (int order = 2; order <= 8; order++) {
                power *= -value;
                sum += power / order;
            }
            return sum;
        }

        private static double FinancialExpm1(double value) {
            if (Math.Abs(value) >= 0.0001) return Math.Exp(value) - 1;
            double sum = value, term = value;
            for (int order = 2; order <= 8; order++) {
                term *= value / order;
                sum += term;
            }
            return sum;
        }
    }
}
