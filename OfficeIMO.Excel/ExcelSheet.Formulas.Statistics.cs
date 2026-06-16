using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateStatisticalFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            if (function == "PERCENTILE.INC" || function == "PERCENTILE.EXC") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count == 0
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double percentile)
                    || percentile < 0d
                    || percentile > 1d
                    || (function == "PERCENTILE.EXC" && (percentile <= 0d || percentile >= 1d))) {
                    return false;
                }

                if (function == "PERCENTILE.EXC") {
                    return TryCalculatePercentileExclusive(values, percentile, out result);
                }

                result = CalculatePercentileInclusive(values, percentile);
                return IsFinite(result);
            }

            if (function == "QUARTILE.INC" || function == "QUARTILE.EXC") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count == 0
                    || !TryGetWholeNumberArgument(tokens[1], out int quartile)
                    || quartile < 0
                    || quartile > 4
                    || (function == "QUARTILE.EXC" && (quartile < 1 || quartile > 3))) {
                    return false;
                }

                if (function == "QUARTILE.EXC") {
                    return TryCalculatePercentileExclusive(values, quartile / 4d, out result);
                }

                result = CalculatePercentileInclusive(values, quartile / 4d);
                return IsFinite(result);
            }

            if (function == "PERCENTRANK.INC" || function == "PERCENTRANK.EXC") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count < 2
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double number)) {
                    return false;
                }

                int? significance = null;
                if (tokens.Count == 3) {
                    if (!TryGetWholeNumberArgument(tokens[2], out int digits) || digits < 1 || digits > 15) {
                        return false;
                    }

                    significance = digits;
                }

                return function == "PERCENTRANK.EXC"
                    ? TryEvaluatePercentRankExclusive(values, number, significance, out result)
                    : TryEvaluatePercentRankInclusive(values, number, significance, out result);
            }

            if (function == "RANK.EQ" || function == "RANK.AVG") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double number)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var values)
                    || values.Count == 0) {
                    return false;
                }

                int order = 0;
                if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out order)) {
                    return false;
                }

                if (!values.Any(value => AreFormulaNumbersEqual(value, number))) {
                    return false;
                }

                int equalCount = values.Count(value => AreFormulaNumbersEqual(value, number));
                int betterCount = order == 0
                    ? values.Count(value => value > number && !AreFormulaNumbersEqual(value, number))
                    : values.Count(value => value < number && !AreFormulaNumbersEqual(value, number));

                double firstRank = betterCount + 1d;
                result = function == "RANK.AVG"
                    ? (firstRank + betterCount + equalCount) / 2d
                    : firstRank;
                return true;
            }

            if (function == "COVAR" || function == "COVARIANCE.P" || function == "COVARIANCE.S") {
                int remainingCellBudget = MaxResolvedFormulaRangeCells;
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var leftValues, ref remainingCellBudget)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var rightValues, ref remainingCellBudget)
                    || !TryCalculateCovariance(leftValues, rightValues, sample: function == "COVARIANCE.S", out result)) {
                    return false;
                }

                return IsFinite(result);
            }

            if (function == "CORREL" || function == "SLOPE" || function == "INTERCEPT" || function == "RSQ") {
                int remainingCellBudget = MaxResolvedFormulaRangeCells;
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var knownY, ref remainingCellBudget)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var knownX, ref remainingCellBudget)
                    || !TryCalculateLinearRegression(knownX, knownY, out double slope, out double intercept, out double correlation)) {
                    return false;
                }

                if (function == "CORREL") {
                    result = correlation;
                } else if (function == "SLOPE") {
                    result = slope;
                } else if (function == "INTERCEPT") {
                    result = intercept;
                } else {
                    result = correlation * correlation;
                }

                return IsFinite(result);
            }

            if (function == "FORECAST.LINEAR") {
                int remainingCellBudget = MaxResolvedFormulaRangeCells;
                if (tokens.Count != 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double x)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var knownY, ref remainingCellBudget)
                    || !TryResolveFormulaArgumentNumbers(tokens[2], out var knownX, ref remainingCellBudget)
                    || !TryCalculateLinearRegression(knownX, knownY, out double slope, out double intercept, out _)) {
                    return false;
                }

                result = intercept + slope * x;
                return IsFinite(result);
            }

            if (function == "SUMXMY2" || function == "SUMX2MY2" || function == "SUMX2PY2") {
                int remainingCellBudget = MaxResolvedFormulaRangeCells;
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var leftValues, ref remainingCellBudget)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var rightValues, ref remainingCellBudget)
                    || leftValues.Count == 0
                    || leftValues.Count != rightValues.Count) {
                    return false;
                }

                result = 0;
                for (int index = 0; index < leftValues.Count; index++) {
                    double left = leftValues[index];
                    double right = rightValues[index];
                    if (function == "SUMXMY2") {
                        double delta = left - right;
                        result += delta * delta;
                    } else if (function == "SUMX2MY2") {
                        result += (left * left) - (right * right);
                    } else {
                        result += (left * left) + (right * right);
                    }
                }

                return IsFinite(result);
            }

            if (!TryResolveFormulaArguments(args, out var arguments) || arguments.Any(value => value.IsUnresolvedFormula)) {
                return false;
            }

            var numbers = arguments.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "AVEDEV" || function == "DEVSQ") {
                if (numbers.Count == 0) {
                    return false;
                }

                double average = numbers.Average();
                if (function == "AVEDEV") {
                    result = numbers.Sum(value => Math.Abs(value - average)) / numbers.Count;
                } else {
                    result = numbers.Sum(value => {
                        double delta = value - average;
                        return delta * delta;
                    });
                }

                return IsFinite(result);
            }

            if (function == "MODE.SNGL" || function == "MODE") {
                return TryCalculateModeSingle(numbers, out result);
            }

            if (function == "GEOMEAN") {
                if (numbers.Count == 0 || numbers.Any(value => value <= 0d)) {
                    return false;
                }

                double averageLog = numbers.Sum(Math.Log) / numbers.Count;
                result = Math.Exp(averageLog);
                return IsFinite(result);
            }

            if (function == "HARMEAN") {
                if (numbers.Count == 0 || numbers.Any(value => value <= 0d)) {
                    return false;
                }

                double reciprocalSum = numbers.Sum(value => 1d / value);
                if (Math.Abs(reciprocalSum) < double.Epsilon) {
                    return false;
                }

                result = numbers.Count / reciprocalSum;
                return IsFinite(result);
            }

            if (function == "VAR.S" || function == "STDEV.S") {
                if (numbers.Count < 2) {
                    return false;
                }

                double variance = CalculateVariance(numbers, sample: true);
                result = function == "STDEV.S" ? Math.Sqrt(variance) : variance;
                return IsFinite(result);
            }

            if (function == "VAR.P" || function == "STDEV.P") {
                if (numbers.Count < 1) {
                    return false;
                }

                double variance = CalculateVariance(numbers, sample: false);
                result = function == "STDEV.P" ? Math.Sqrt(variance) : variance;
                return IsFinite(result);
            }

            return false;
        }

        private static bool TryCalculateModeSingle(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count == 0) {
                return false;
            }

            var groups = numbers
                .Select((value, index) => new { Value = value, Index = index })
                .GroupBy(item => item.Value)
                .Select(group => new { Value = group.Key, Count = group.Count(), FirstIndex = group.Min(item => item.Index) })
                .OrderByDescending(group => group.Count)
                .ThenBy(group => group.FirstIndex)
                .ToList();

            if (groups.Count == 0 || groups[0].Count < 2) {
                return false;
            }

            result = groups[0].Value;
            return IsFinite(result);
        }

        private static double CalculateVariance(IReadOnlyList<double> numbers, bool sample) {
            double average = numbers.Average();
            double sumSquares = numbers.Sum(value => {
                double delta = value - average;
                return delta * delta;
            });
            return sumSquares / (sample ? numbers.Count - 1 : numbers.Count);
        }

        private static bool TryCalculateCovariance(IReadOnlyList<double> leftValues, IReadOnlyList<double> rightValues, bool sample, out double result) {
            result = 0;
            if (leftValues.Count != rightValues.Count || leftValues.Count == 0 || (sample && leftValues.Count < 2)) {
                return false;
            }

            double leftAverage = leftValues.Average();
            double rightAverage = rightValues.Average();
            double sumProducts = 0;
            for (int index = 0; index < leftValues.Count; index++) {
                sumProducts += (leftValues[index] - leftAverage) * (rightValues[index] - rightAverage);
            }

            result = sumProducts / (sample ? leftValues.Count - 1 : leftValues.Count);
            return IsFinite(result);
        }

        private static bool TryCalculateLinearRegression(
            IReadOnlyList<double> knownX,
            IReadOnlyList<double> knownY,
            out double slope,
            out double intercept,
            out double correlation) {
            slope = 0;
            intercept = 0;
            correlation = 0;
            if (knownX.Count != knownY.Count || knownX.Count < 2) {
                return false;
            }

            double averageX = knownX.Average();
            double averageY = knownY.Average();
            double sumXy = 0;
            double sumXx = 0;
            double sumYy = 0;
            for (int index = 0; index < knownX.Count; index++) {
                double xDelta = knownX[index] - averageX;
                double yDelta = knownY[index] - averageY;
                sumXy += xDelta * yDelta;
                sumXx += xDelta * xDelta;
                sumYy += yDelta * yDelta;
            }

            if (Math.Abs(sumXx) < double.Epsilon || Math.Abs(sumYy) < double.Epsilon) {
                return false;
            }

            slope = sumXy / sumXx;
            intercept = averageY - slope * averageX;
            correlation = sumXy / Math.Sqrt(sumXx * sumYy);
            return IsFinite(slope) && IsFinite(intercept) && IsFinite(correlation);
        }

        private static bool AreFormulaNumbersEqual(double left, double right) {
            return Math.Abs(left - right) < 0.0000001d;
        }

        private static bool TryEvaluatePercentRankInclusive(IReadOnlyList<double> numbers, double number, int? significance, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            if (number < sorted[0] || number > sorted[sorted.Count - 1]) {
                return false;
            }

            double denominator = sorted.Count - 1d;
            for (int index = 0; index < sorted.Count; index++) {
                if (AreFormulaNumbersEqual(sorted[index], number)) {
                    result = index / denominator;
                    if (significance.HasValue) {
                        result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                    }

                    return IsFinite(result);
                }
            }

            for (int index = 0; index < sorted.Count - 1; index++) {
                double lower = sorted[index];
                double upper = sorted[index + 1];
                if (number <= lower || number >= upper) {
                    continue;
                }

                double fraction = (number - lower) / (upper - lower);
                result = (index + fraction) / denominator;
                if (significance.HasValue) {
                    result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                }

                return IsFinite(result);
            }

            return false;
        }

        private static bool TryEvaluatePercentRankExclusive(IReadOnlyList<double> numbers, double number, int? significance, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            if (number < sorted[0] || number > sorted[sorted.Count - 1]) {
                return false;
            }

            double denominator = sorted.Count + 1d;
            for (int index = 0; index < sorted.Count; index++) {
                if (AreFormulaNumbersEqual(sorted[index], number)) {
                    result = (index + 1d) / denominator;
                    if (significance.HasValue) {
                        result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                    }

                    return IsFinite(result);
                }
            }

            for (int index = 0; index < sorted.Count - 1; index++) {
                double lower = sorted[index];
                double upper = sorted[index + 1];
                if (number <= lower || number >= upper) {
                    continue;
                }

                double fraction = (number - lower) / (upper - lower);
                result = (index + 1d + fraction) / denominator;
                if (significance.HasValue) {
                    result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                }

                return IsFinite(result);
            }

            return false;
        }

        private static bool TryCalculatePercentileExclusive(IReadOnlyList<double> numbers, double percentile, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            double rank = percentile * (sorted.Count + 1d);
            if (rank <= 1d || rank >= sorted.Count) {
                return false;
            }

            int lowerRank = (int)Math.Floor(rank);
            int upperRank = (int)Math.Ceiling(rank);
            if (lowerRank == upperRank) {
                result = sorted[lowerRank - 1];
                return IsFinite(result);
            }

            double fraction = rank - lowerRank;
            double lower = sorted[lowerRank - 1];
            double upper = sorted[upperRank - 1];
            result = lower + (upper - lower) * fraction;
            return IsFinite(result);
        }

        private static double CalculatePercentileInclusive(IReadOnlyList<double> numbers, double percentile) {
            var sorted = numbers.OrderBy(value => value).ToList();
            if (sorted.Count == 1) {
                return sorted[0];
            }

            double rank = (sorted.Count - 1) * percentile;
            int lower = (int)Math.Floor(rank);
            int upper = (int)Math.Ceiling(rank);
            if (lower == upper) {
                return sorted[lower];
            }

            double fraction = rank - lower;
            return sorted[lower] + (sorted[upper] - sorted[lower]) * fraction;
        }

    }
}
