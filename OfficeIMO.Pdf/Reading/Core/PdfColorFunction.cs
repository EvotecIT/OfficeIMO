namespace OfficeIMO.Pdf;

/// <summary>Immutable, bounded PDF function used by color spaces and shading projection.</summary>
internal sealed class PdfColorFunction {
    private readonly double[] _domain;
    private readonly double[]? _range;
    private readonly Func<double[], double[]?> _evaluateCore;
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<double> _breakpoints;
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<double> _discontinuities;

    internal PdfColorFunction(
        int inputCount,
        int outputCount,
        double[] domain,
        double[]? range,
        Func<double[], double[]?> evaluateCore,
        IReadOnlyList<double>? breakpoints = null,
        IReadOnlyList<double>? discontinuities = null,
        int evaluationCost = 0) {
        InputCount = inputCount;
        OutputCount = outputCount;
        _domain = (double[])domain.Clone();
        _range = range == null ? null : (double[])range.Clone();
        _evaluateCore = evaluateCore;
        EvaluationCost = Math.Max(0, evaluationCost);

        double[] points = breakpoints == null
            ? Array.Empty<double>()
            : breakpoints.Where(static value => !double.IsNaN(value) && !double.IsInfinity(value)).Distinct().OrderBy(static value => value).ToArray();
        _breakpoints = Array.AsReadOnly(points);
        double[] edges = discontinuities == null
            ? Array.Empty<double>()
            : discontinuities.Where(static value => !double.IsNaN(value) && !double.IsInfinity(value)).Distinct().OrderBy(static value => value).ToArray();
        _discontinuities = Array.AsReadOnly(edges);
    }

    internal int InputCount { get; }

    internal int OutputCount { get; }

    /// <summary>Worst-case bounded work units for one non-trivial evaluation; zero for constant-cost functions.</summary>
    internal int EvaluationCost { get; }

    internal IReadOnlyList<double> Domain => _domain;

    internal bool HasRange => _range != null;

    internal IReadOnlyList<double> Range => _range ?? Array.Empty<double>();

    /// <summary>Authored input positions worth retaining when a one-input function is projected to gradient stops.</summary>
    internal IReadOnlyList<double> Breakpoints => _breakpoints;

    /// <summary>Authored one-input boundaries whose left and right values may differ.</summary>
    internal IReadOnlyList<double> Discontinuities => _discontinuities;

    internal double[]? Evaluate(IReadOnlyList<double> values) {
        if (values == null || values.Count < InputCount) return null;

        double[]? clipped = null;
        double[]? sourceArray = values as double[];
        for (int index = 0; index < InputCount; index++) {
            double value = values[index];
            if (!IsFinite(value)) return null;
            double bounded = Clamp(value, _domain[index * 2], _domain[index * 2 + 1]);
            if (bounded == value && clipped == null) continue;
            if (clipped == null) {
                clipped = new double[InputCount];
                for (int copy = 0; copy < index; copy++) clipped[copy] = values[copy];
            }
            clipped[index] = bounded;
        }

        double[] input = clipped ?? (sourceArray != null && sourceArray.Length == InputCount ? sourceArray : values.Take(InputCount).ToArray());
        double[]? result = _evaluateCore(input);
        if (result == null || result.Length != OutputCount || result.Any(static value => !IsFinite(value))) return null;
        if (_range != null) {
            for (int index = 0; index < result.Length; index++) {
                result[index] = Clamp(result[index], _range[index * 2], _range[index * 2 + 1]);
            }
        }

        return result;
    }

    internal static double Interpolate(double value, double sourceStart, double sourceEnd, double targetStart, double targetEnd) {
        if (sourceStart == sourceEnd) return targetStart;
        if (value == sourceStart) return targetStart;
        if (value == sourceEnd) return targetEnd;

        double sourceSpan = sourceEnd - sourceStart;
        double fraction = IsFinite(sourceSpan)
            ? (value - sourceStart) / sourceSpan
            : (value * 0.5D - sourceStart * 0.5D) / (sourceEnd * 0.5D - sourceStart * 0.5D);
        if (!IsFinite(fraction)) return double.NaN;

        double result = targetStart * (1D - fraction) + targetEnd * fraction;
        return IsFinite(result) ? result : double.NaN;
    }

    internal static double Clamp(double value, double minimum, double maximum) =>
        value < minimum ? minimum : value > maximum ? maximum : value;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
