using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateFinancialFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);

            if (function == "NPV") {
                return TryEvaluateNpv(tokens, out result);
            }

            if (!TryResolveFormulaOrNumericArguments(tokens, out var numbers)) {
                return false;
            }

            switch (function) {
                case "PMT":
                    return TryEvaluatePmt(numbers, out result);
                case "PV":
                    return TryEvaluatePv(numbers, out result);
                case "FV":
                    return TryEvaluateFv(numbers, out result);
                case "NPER":
                    return TryEvaluateNper(numbers, out result);
                default:
                    return false;
            }
        }

        private static bool TryEvaluatePmt(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double presentValue = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type) || Math.Abs(periods) < double.Epsilon) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(presentValue + futureValue) / periods;
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor) || Math.Abs(factor - 1d) < double.Epsilon) {
                return false;
            }

            result = -(rate * (futureValue + presentValue * factor)) / ((1d + rate * type) * (factor - 1d));
            return IsFinite(result);
        }

        private static bool TryEvaluatePv(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double payment = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(futureValue + payment * periods);
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor) || Math.Abs(factor) < double.Epsilon) {
                return false;
            }

            result = -(futureValue + payment * (1d + rate * type) * ((factor - 1d) / rate)) / factor;
            return IsFinite(result);
        }

        private static bool TryEvaluateFv(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double payment = numbers[2];
            double presentValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(presentValue + payment * periods);
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor)) {
                return false;
            }

            result = -(presentValue * factor + payment * (1d + rate * type) * ((factor - 1d) / rate));
            return IsFinite(result);
        }

        private static bool TryEvaluateNper(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double payment = numbers[1];
            double presentValue = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                if (Math.Abs(payment) < double.Epsilon) {
                    return false;
                }

                result = -(presentValue + futureValue) / payment;
                return IsFinite(result);
            }

            double adjustedPayment = payment * (1d + rate * type);
            double numerator = adjustedPayment - futureValue * rate;
            double denominator = presentValue * rate + adjustedPayment;
            if (Math.Abs(denominator) < double.Epsilon || 1d + rate <= 0d) {
                return false;
            }

            double ratio = numerator / denominator;
            if (ratio <= 0d) {
                return false;
            }

            result = Math.Log(ratio) / Math.Log(1d + rate);
            return IsFinite(result);
        }

        private bool TryEvaluateNpv(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2 || !TryEvaluateFormulaOrNumeric(tokens[0], out double rate) || 1d + rate == 0d) {
                return false;
            }

            int period = 1;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            for (int index = 1; index < tokens.Count; index++) {
                if (!TryResolveFormulaArgumentNumbers(tokens[index], out var values, ref remainingCellBudget) || values.Count == 0) {
                    return false;
                }

                foreach (double value in values) {
                    double divisor = Math.Pow(1d + rate, period);
                    if (!IsFinite(divisor) || Math.Abs(divisor) < double.Epsilon) {
                        return false;
                    }

                    result += value / divisor;
                    period++;
                }
            }

            return IsFinite(result);
        }

        private static bool TryGetPaymentType(IReadOnlyList<double> numbers, int index, out int type) {
            type = 0;
            if (numbers.Count <= index) {
                return true;
            }

            if (!TryGetWholeNumber(numbers[index], out type)) {
                return false;
            }

            return type == 0 || type == 1;
        }

        private static bool IsFinite(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static bool TryEvaluateMRound(double number, double multiple, out double result) {
            result = 0;
            if (!IsFinite(number) || !IsFinite(multiple)) {
                return false;
            }

            if (Math.Abs(multiple) < double.Epsilon) {
                result = 0;
                return true;
            }

            if (Math.Sign(number) != 0 && Math.Sign(multiple) != 0 && Math.Sign(number) != Math.Sign(multiple)) {
                return false;
            }

            result = Math.Round(number / multiple, 0, MidpointRounding.AwayFromZero) * multiple;
            return IsFinite(result);
        }

        private static bool TryEvaluateMathRoundFunction(string function, IReadOnlyList<double> numbers, out double result) {
            result = 0;
            double number = numbers[0];
            double significance = numbers.Count >= 2 ? Math.Abs(numbers[1]) : 1d;
            double mode = numbers.Count >= 3 ? numbers[2] : 0d;
            if (!IsFinite(number) || !IsFinite(significance) || !IsFinite(mode)) {
                return false;
            }

            if (Math.Abs(significance) < double.Epsilon) {
                result = 0;
                return true;
            }

            double quotient = number / significance;
            if (number >= 0) {
                result = (function == "CEILING.MATH" ? Math.Ceiling(quotient) : Math.Floor(quotient)) * significance;
                return IsFinite(result);
            }

            bool useAwayFromZeroForNegative = function == "CEILING.MATH"
                ? Math.Abs(mode) >= double.Epsilon
                : Math.Abs(mode) < double.Epsilon;
            result = (useAwayFromZeroForNegative ? Math.Floor(quotient) : Math.Ceiling(quotient)) * significance;
            return IsFinite(result);
        }

    }
}
