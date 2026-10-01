namespace OfficeIMO.Excel {
    /// <summary>Single-pass aggregate state shared by pivot groups and their totals.</summary>
    internal sealed class ExcelPivotAggregateAccumulator {
        private long _nonEmptyCount;
        private long _numericCount;
        private double _sum;
        private double _product = 1;
        private double _minimum = double.PositiveInfinity;
        private double _maximum = double.NegativeInfinity;
        private double _mean;
        private double _squaredDeviations;
        private string? _error;

        /// <summary>Adds a typed worksheet cache value; text and Boolean values are not numeric.</summary>
        internal void Add(object? value, bool isError = false) {
            if (value == null) return;
            _nonEmptyCount++;
            if (isError) {
                _error ??= value.ToString() ?? "#VALUE!";
                return;
            }
            if (value is not double number) return;
            _numericCount++;
            _sum += number;
            _product *= number;
            _minimum = Math.Min(_minimum, number);
            _maximum = Math.Max(_maximum, number);
            double difference = number - _mean;
            _mean += difference / _numericCount;
            _squaredDeviations += difference * (number - _mean);
        }

        /// <summary>Combines disjoint group states without averaging averages or summing variances.</summary>
        internal void Merge(ExcelPivotAggregateAccumulator other) {
            if (ReferenceEquals(this, other)) throw new ArgumentException("A pivot aggregate cannot merge itself.", nameof(other));
            _nonEmptyCount += other._nonEmptyCount;
            _error ??= other._error;
            if (other._numericCount == 0) return;
            if (_numericCount == 0) {
                _numericCount = other._numericCount;
                _sum = other._sum;
                _product = other._product;
                _minimum = other._minimum;
                _maximum = other._maximum;
                _mean = other._mean;
                _squaredDeviations = other._squaredDeviations;
                return;
            }
            long count = _numericCount + other._numericCount;
            double difference = other._mean - _mean;
            _squaredDeviations += other._squaredDeviations
                + difference * difference * ((double)_numericCount / count) * other._numericCount;
            _mean += difference * ((double)other._numericCount / count);
            _numericCount = count;
            _sum += other._sum;
            _product *= other._product;
            _minimum = Math.Min(_minimum, other._minimum);
            _maximum = Math.Max(_maximum, other._maximum);
        }

        /// <summary>Returns a typed numeric or error result for one pivot aggregation mode.</summary>
        internal ExcelCellData GetValue(ExcelPivotDataFunction function) {
            if ((uint)function > (uint)ExcelPivotDataFunction.VarianceP)
                throw new ArgumentOutOfRangeException(nameof(function));
            if (function == ExcelPivotDataFunction.Count) return Number(_nonEmptyCount);
            if (function == ExcelPivotDataFunction.CountNumbers) return Number(_numericCount);
            if (_error != null) return Error(_error);
            if (_numericCount == 0) {
                bool requiresNumbers = function == ExcelPivotDataFunction.Average
                    || function == ExcelPivotDataFunction.StandardDeviation || function == ExcelPivotDataFunction.StandardDeviationP
                    || function == ExcelPivotDataFunction.Variance || function == ExcelPivotDataFunction.VarianceP;
                return _nonEmptyCount != 0 && requiresNumbers ? Error("#DIV/0!") : Number(0);
            }
            bool sample = function == ExcelPivotDataFunction.StandardDeviation || function == ExcelPivotDataFunction.Variance;
            if (sample && _numericCount == 1) return Error("#DIV/0!");
            double variance = Math.Max(0, _squaredDeviations / (_numericCount - (sample ? 1 : 0)));
            double result = function switch {
                ExcelPivotDataFunction.Sum => _sum,
                ExcelPivotDataFunction.Average => _mean,
                ExcelPivotDataFunction.Minimum => _minimum,
                ExcelPivotDataFunction.Maximum => _maximum,
                ExcelPivotDataFunction.Product => _product,
                ExcelPivotDataFunction.StandardDeviation or ExcelPivotDataFunction.StandardDeviationP => Math.Sqrt(variance),
                ExcelPivotDataFunction.Variance or ExcelPivotDataFunction.VarianceP => variance,
                _ => throw new ArgumentOutOfRangeException(nameof(function))
            };
            return double.IsNaN(result) || double.IsInfinity(result) ? Error("#NUM!") : Number(result);
        }

        private static ExcelCellData Number(double value) => new(ExcelCellDataKind.Number, value);
        private static ExcelCellData Error(string value) => new(ExcelCellDataKind.Error, value, cachedText: value);
    }
}
