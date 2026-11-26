namespace OfficeIMO.Pdf;

/// <summary>
/// Single-pass minimum, maximum and sum of doubles with the same results as LINQ's <c>Min</c>, <c>Max</c>,
/// <c>Sum</c> and <c>Average</c>, NaN handling included, without an enumerator per call.
/// </summary>
internal struct PdfDoubleAggregate {
    private double _minimum;
    private double _maximum;
    private bool _minimumIsNaN;
    private int _count;

    /// <summary>Values added so far.</summary>
    internal readonly int Count => _count;

    /// <summary>Running sum, starting from zero as <c>Enumerable.Sum</c> does.</summary>
    internal double Sum { readonly get; private set; }

    /// <summary>Running sum, starting from the first value as <c>Enumerable.Average</c> does.</summary>
    private double _averageSum;

    internal void Add(double value) {
        if (_count == 0) {
            _minimum = value;
            _minimumIsNaN = double.IsNaN(value);
            _maximum = value;
            _averageSum = value;
        } else {
            // Enumerable.Min returns NaN as soon as it sees one.
            if (!_minimumIsNaN) {
                if (value < _minimum) _minimum = value;
                else if (double.IsNaN(value)) { _minimum = value; _minimumIsNaN = true; }
            }
            // Enumerable.Max skips NaN unless every value is NaN.
            if (double.IsNaN(_maximum) || value > _maximum) _maximum = value;
            _averageSum += value;
        }
        Sum += value;
        _count++;
    }

    internal readonly double Minimum => _count == 0 ? throw new InvalidOperationException("Sequence contains no elements.") : _minimum;
    internal readonly double Maximum => _count == 0 ? throw new InvalidOperationException("Sequence contains no elements.") : _maximum;
    internal readonly double Average => _count == 0 ? throw new InvalidOperationException("Sequence contains no elements.") : _averageSum / _count;
}
