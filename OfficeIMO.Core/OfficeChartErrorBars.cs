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
                double[] finite = values.Where(item => !double.IsNaN(item) && !double.IsInfinity(item)).ToArray();
                if (finite.Length < 2) return false;
                double mean = finite.Average();
                double squares = finite.Sum(item => (item - mean) * (item - mean));
                if (ValueType == OfficeChartErrorBarValueType.StandardError) {
                    plus = minus = Math.Sqrt(squares / (finite.Length * (finite.Length - 1D)));
                } else {
                    center = mean;
                    low = high = mean;
                    plus = minus = Math.Abs(Amount) * Math.Sqrt(squares / (finite.Length - 1D));
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
}
