namespace OfficeIMO.Drawing;

/// <summary>How an error bar's length is calculated.</summary>
internal enum OfficeChartErrorBarValueType {
    Fixed,
    Percentage,
    StandardDeviation,
    StandardError,
    Custom
}

/// <summary>Which side of a value an error bar extends to.</summary>
internal enum OfficeChartErrorBarDirection {
    Both,
    Plus,
    Minus
}

/// <summary>Native error bars read from a chart; drawn by the static chart renderer.</summary>
internal sealed class OfficeChartErrorBars {
    internal OfficeChartErrorBars(OfficeChartErrorBarValueType valueType, double amount, OfficeChartErrorBarDirection direction,
        bool endCaps, bool alongCategoryAxis, IReadOnlyList<double>? plus = null, IReadOnlyList<double>? minus = null,
        OfficeColor? color = null, double? width = null, OfficeStrokeDashStyle? dashStyle = null) {
        ValueType = valueType;
        Amount = amount;
        Direction = direction;
        EndCaps = endCaps;
        AlongCategoryAxis = alongCategoryAxis;
        Plus = plus;
        Minus = minus;
        Color = color;
        Width = width;
        DashStyle = dashStyle;
    }

    internal OfficeChartErrorBarValueType ValueType { get; }
    internal double Amount { get; }
    internal OfficeChartErrorBarDirection Direction { get; }
    internal bool EndCaps { get; }
    /// <summary>Whether the bars run along the x (category or horizontal value) axis of a scatter chart.</summary>
    internal bool AlongCategoryAxis { get; }
    internal IReadOnlyList<double>? Plus { get; }
    internal IReadOnlyList<double>? Minus { get; }
    internal OfficeColor? Color { get; }
    internal double? Width { get; }
    internal OfficeStrokeDashStyle? DashStyle { get; }

    /// <summary>
    /// Resolves the drawn span for one point, as Office does: percentage and fixed bars surround the value,
    /// standard error uses the series' standard error, and standard deviation bars are centred on the series mean.
    /// </summary>
    internal bool TryGetSpan(IReadOnlyList<double> values, int index, double value, out double center, out double low, out double high) {
        center = value;
        low = high = value;
        double plus, minus;
        switch (ValueType) {
            case OfficeChartErrorBarValueType.Fixed:
                plus = minus = Math.Abs(Amount);
                break;
            case OfficeChartErrorBarValueType.Percentage:
                plus = minus = Math.Abs(value * Amount / 100D);
                break;
            case OfficeChartErrorBarValueType.StandardError:
            case OfficeChartErrorBarValueType.StandardDeviation:
                SeriesStatistics statistics = GetStatistics(values);
                if (statistics.Count < 2) return false;
                if (ValueType == OfficeChartErrorBarValueType.StandardError) {
                    plus = minus = Math.Sqrt(statistics.Squares / (statistics.Count * (statistics.Count - 1D)));
                } else {
                    center = statistics.Mean;
                    low = high = statistics.Mean;
                    plus = minus = Math.Abs(Amount) * Math.Sqrt(statistics.Squares / (statistics.Count - 1D));
                }
                break;
            default:
                plus = Plus != null && index < Plus.Count ? Plus[index] : 0D;
                minus = Minus != null && index < Minus.Count ? Minus[index] : 0D;
                break;
        }
        if (double.IsNaN(plus) || double.IsInfinity(plus) || double.IsNaN(minus) || double.IsInfinity(minus)) return false;
        if (Direction != OfficeChartErrorBarDirection.Minus) high = center + plus;
        if (Direction != OfficeChartErrorBarDirection.Plus) low = center - minus;
        return true;
    }

    // Drawing and axis scaling ask for every point's span with the same value list, so the series statistics are
    // kept for the last list seen instead of being recomputed per point.
    private SeriesStatistics? _statistics;

    private SeriesStatistics GetStatistics(IReadOnlyList<double> values) {
        SeriesStatistics? cached = _statistics;
        if (cached != null && ReferenceEquals(cached.Values, values) && cached.Values.Count == values.Count) return cached;
        int count = 0;
        double sum = 0D;
        foreach (double value in values) {
            if (double.IsNaN(value) || double.IsInfinity(value)) continue;
            count++;
            sum += value;
        }
        double mean = count == 0 ? 0D : sum / count;
        double squares = 0D;
        if (count > 0) {
            foreach (double value in values) {
                if (double.IsNaN(value) || double.IsInfinity(value)) continue;
                squares += (value - mean) * (value - mean);
            }
        }
        var statistics = new SeriesStatistics(values, count, mean, squares);
        _statistics = statistics;
        return statistics;
    }

    private sealed class SeriesStatistics {
        internal SeriesStatistics(IReadOnlyList<double> values, int count, double mean, double squares) {
            Values = values;
            Count = count;
            Mean = mean;
            Squares = squares;
        }

        internal IReadOnlyList<double> Values { get; }
        internal int Count { get; }
        internal double Mean { get; }
        internal double Squares { get; }
    }
}
