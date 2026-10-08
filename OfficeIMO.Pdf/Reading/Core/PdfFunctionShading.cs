using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>A two-dimensional PDF color field, retained without finite-stop approximation.</summary>
internal sealed class PdfFunctionShading {
    private readonly PdfColorFunction[] _functions;
    private readonly PdfPageColorSpace _colorSpace;
    private readonly OfficeIccRenderingIntent _intent;
    private readonly PdfOutputIntentColorTransform? _outputIntent;
    private readonly Func<int, bool> _consumeWork;
    private readonly CancellationToken _cancellationToken;
    private readonly OfficeTransform _inverse;
    private readonly double[] _domain;
    private readonly double[]? _bounds;
    private readonly int _evaluationCost;

    internal PdfFunctionShading(PdfColorFunction[] functions, PdfPageColorSpace colorSpace,
        double[] domain, OfficeTransform matrix, double[]? bounds,
        OfficeIccRenderingIntent intent, PdfOutputIntentColorTransform? outputIntent,
        Func<int, bool> consumeWork, CancellationToken cancellationToken, double rasterScale = 1D, Action<long>? chargePixels = null, PdfDictionary? sourceDictionary = null) {
        SourceDictionary = sourceDictionary;
        RasterScale = rasterScale; ChargePixels = chargePixels;
        _functions = (PdfColorFunction[])functions.Clone();
        _colorSpace = colorSpace; _domain = (double[])domain.Clone();
        _inverse = matrix.Invert(); _bounds = bounds == null ? null : (double[])bounds.Clone();
        _intent = intent; _outputIntent = outputIntent; _consumeWork = consumeWork;
        _cancellationToken = cancellationToken;
        long cost = 0;
        foreach (var function in functions) cost += Math.Max(1, function.EvaluationCost);
        _evaluationCost = (int)Math.Min(int.MaxValue, cost);
    }

    internal PdfDictionary? SourceDictionary { get; }

    internal CancellationToken CancellationToken => _cancellationToken;
    internal double RasterScale { get; }
    internal Action<long>? ChargePixels { get; }
    internal int ComponentCount => _colorSpace.ComponentCount;

    // Scratch buffers belong to one renderer invocation; no per-pixel allocation
    // or mutable shared input survives across callers.
    internal bool TrySample(double x, double y, double[] input, double[] components, out OfficeColor color) {
        _cancellationToken.ThrowIfCancellationRequested();
        color = OfficeColor.Transparent;
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y) ||
            input.Length < 2 || components.Length < ComponentCount) return false;
        if (_bounds != null && (x < _bounds[0] || y < _bounds[1] || x > _bounds[2] || y > _bounds[3])) return true;
        var point = _inverse.TransformPoint(new OfficePoint(x, y));
        if (point.X < _domain[0] || point.X > _domain[1] || point.Y < _domain[2] || point.Y > _domain[3]) return true;
        input[0] = point.X; input[1] = point.Y;
        if (!_consumeWork(_evaluationCost)) return false;
        for (int index = 0; index < _functions.Length; index++) {
            if (!_functions[index].TryEvaluate(input, components, _functions.Length == 1 ? 0 : index)) return false;
        }
        if (_outputIntent != null && _outputIntent.TryApplyDirect(_colorSpace, components, _intent, out color)) return true;
        if (!_colorSpace.TryConvertColor(components, _intent, out color)) return false;
        if (_outputIntent != null) color = _outputIntent.Apply(color, _intent);
        return true;
    }
}
