namespace OfficeIMO.Pdf;

/// <summary>Samples the colour components of a PDF shading at a point in shading space.</summary>
internal interface IPdfShadingFieldSampler {
    /// <summary>Bounded work units charged for one sample.</summary>
    int EvaluationCost { get; }
    int GetEvaluationCost(double x, double y);

    /// <summary>Writes the colour components at a point; returns false where the shading paints nothing.</summary>
    bool TrySampleComponents(double x, double y, double[] input, double[] components);

    /// <summary>Shading-space bounds outside which nothing paints, or null for an unbounded field.</summary>
    double[]? PaintBounds { get; }
}

/// <summary>Evaluates a parametric shading's colour functions at t.</summary>
internal sealed class PdfShadingColorFunctions {
    private readonly PdfColorFunction[] _functions;

    internal PdfShadingColorFunctions(PdfColorFunction[] functions, double t0, double t1) {
        _functions = functions;
        T0 = t0;
        T1 = t1;
        long cost = 0;
        foreach (PdfColorFunction function in functions) cost += Math.Max(1, function.EvaluationCost);
        EvaluationCost = (int)Math.Min(int.MaxValue, cost);
    }

    internal double T0 { get; }
    internal double T1 { get; }
    internal int EvaluationCost { get; }

    internal bool TryEvaluate(double t, double[] input, double[] components) {
        input[0] = t;
        for (int index = 0; index < _functions.Length; index++) {
            if (!_functions[index].TryEvaluate(input, components, _functions.Length == 1 ? 0 : index)) return false;
        }
        return true;
    }
}

/// <summary>Type 2 axial shading with PDF Extend semantics (ISO 32000-1, 8.7.4.5.3).</summary>
internal sealed class PdfAxialShadingSampler : IPdfShadingFieldSampler {
    private readonly double _x0, _y0, _dx, _dy, _lengthSquared;
    private readonly bool _extendStart, _extendEnd;
    private readonly PdfShadingColorFunctions _functions;

    internal PdfAxialShadingSampler(double x0, double y0, double x1, double y1, bool extendStart, bool extendEnd, PdfShadingColorFunctions functions) {
        _x0 = x0; _y0 = y0; _dx = x1 - x0; _dy = y1 - y0;
        _lengthSquared = _dx * _dx + _dy * _dy;
        _extendStart = extendStart; _extendEnd = extendEnd;
        _functions = functions;
    }

    public int EvaluationCost => _functions.EvaluationCost;
    public int GetEvaluationCost(double x, double y) => EvaluationCost;

    public double[]? PaintBounds => null;

    public bool TrySampleComponents(double x, double y, double[] input, double[] components) {
        if (_lengthSquared <= 0D) return false;
        double s = ((x - _x0) * _dx + (y - _y0) * _dy) / _lengthSquared;
        if (s < 0D) {
            if (!_extendStart) return false;
            s = 0D;
        } else if (s > 1D) {
            if (!_extendEnd) return false;
            s = 1D;
        }
        return _functions.TryEvaluate(_functions.T0 + s * (_functions.T1 - _functions.T0), input, components);
    }
}

/// <summary>
/// Type 3 radial shading: the colour of the largest s whose circle contains the point and whose radius is non-negative,
/// with s limited to [0, 1] unless the corresponding end extends (ISO 32000-1, 8.7.4.5.4).
/// </summary>
internal sealed class PdfRadialShadingSampler : IPdfShadingFieldSampler {
    private readonly double _x0, _y0, _r0, _cdx, _cdy, _dr, _a;
    private readonly bool _pointEnd;
    private readonly bool _extendStart, _extendEnd;
    private readonly PdfShadingColorFunctions _functions;

    internal PdfRadialShadingSampler(double x0, double y0, double r0, double x1, double y1, double r1,
        bool extendStart, bool extendEnd, PdfShadingColorFunctions functions) {
        // Solve from a zero-radius end to avoid subtracting almost equal large squared radii.
        _pointEnd = r1 == 0D;
        _x0 = _pointEnd ? x1 : x0; _y0 = _pointEnd ? y1 : y0; _r0 = _pointEnd ? r1 : r0;
        _cdx = (x1 - x0) * (_pointEnd ? -1D : 1D);
        _cdy = (y1 - y0) * (_pointEnd ? -1D : 1D);
        _dr = (r1 - r0) * (_pointEnd ? -1D : 1D);
        _a = _cdx * _cdx + _cdy * _cdy - _dr * _dr;
        _extendStart = extendStart; _extendEnd = extendEnd;
        _functions = functions;
    }

    public int EvaluationCost => _functions.EvaluationCost;
    public int GetEvaluationCost(double x, double y) => EvaluationCost;

    public double[]? PaintBounds => null;

    public bool TrySampleComponents(double x, double y, double[] input, double[] components) {
        double px = x - _x0, py = y - _y0;
        double b = px * _cdx + py * _cdy + _r0 * _dr;
        double c = px * px + py * py - _r0 * _r0;
        double s;
        if (_a == 0D) {
            if (b == 0D) return false;
            s = c / (2D * b);
            if (!IsValid(s)) return false;
        } else {
            double discriminant = b * b - _a * c;
            if (discriminant < 0D) return false;
            double root = Math.Sqrt(discriminant);
            double q = b + (b >= 0D ? root : -root);
            double high = q / _a, low = q == 0D ? 0D : c / q;
            if (high < low) (high, low) = (low, high);
            double preferred = _pointEnd ? low : high, other = _pointEnd ? high : low;
            if (IsValid(preferred)) s = preferred;
            else if (IsValid(other)) s = other;
            else return false;
        }
        if (_pointEnd) s = 1D - s;
        s = Math.Max(0D, Math.Min(1D, s));
        return _functions.TryEvaluate(_functions.T0 + s * (_functions.T1 - _functions.T0), input, components);
    }

    private bool IsValid(double s) =>
        !double.IsNaN(s) && _r0 + s * _dr >= 0D &&
        (s >= 0D || (_pointEnd ? _extendEnd : _extendStart)) && (s <= 1D || (_pointEnd ? _extendStart : _extendEnd));
}

/// <summary>
/// Triangle field for mesh shadings (types 4 to 7). Later triangles paint over earlier ones, matching the stream order;
/// a uniform grid bounds the per-sample search.
/// </summary>
internal sealed class PdfMeshShadingSampler : IPdfShadingFieldSampler {
    private const long MaximumGridEntries = 8_000_000L;
    private readonly double[] _x, _y, _values;
    private readonly int[] _triangles;
    private readonly int _valueCount;
    private readonly PdfShadingColorFunctions? _functions;
    private readonly double _left, _top, _cellWidth, _cellHeight;
    private readonly int _columns, _rows;
    private readonly int[][] _cells;

    internal PdfMeshShadingSampler(List<double> x, List<double> y, List<double> values, int valueCount, List<int> triangles,
        PdfShadingColorFunctions? functions) {
        _x = x.ToArray(); _y = y.ToArray(); _values = values.ToArray();
        _valueCount = valueCount;
        _triangles = triangles.ToArray();
        _functions = functions;
        double left = double.PositiveInfinity, top = double.PositiveInfinity, right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (double value in _x) { left = Math.Min(left, value); right = Math.Max(right, value); }
        foreach (double value in _y) { top = Math.Min(top, value); bottom = Math.Max(bottom, value); }
        int triangleCount = _triangles.Length / 3;
        int side = Math.Max(1, Math.Min(256, (int)Math.Sqrt(Math.Max(1, triangleCount))));
        _columns = side; _rows = side;
        _left = left; _top = top;
        PaintBounds = new[] { left, top, right, bottom };
        _cellWidth = Math.Max((right - left) / _columns, 1e-9D);
        _cellHeight = Math.Max((bottom - top) / _rows, 1e-9D);
        var counts = new int[_columns * _rows];
        long entries = 0L;
        for (int pass = 0; pass < 2; pass++) {
            for (int triangle = 0; triangle < triangleCount; triangle++) {
                CellRange(triangle, out int c0, out int r0, out int c1, out int r1);
                for (int row = r0; row <= r1; row++) {
                    for (int column = c0; column <= c1; column++) {
                        int cell = row * _columns + column;
                        if (pass == 0) {
                            counts[cell]++;
                            if (++entries > MaximumGridEntries) throw new InvalidDataException("The mesh shading exceeds the sampling index budget.");
                        } else {
                            _cells![cell][--counts[cell]] = triangle;
                        }
                    }
                }
            }
            if (pass == 0) {
                _cells = new int[counts.Length][];
                for (int cell = 0; cell < counts.Length; cell++) _cells[cell] = new int[counts[cell]];
            }
        }
        _cells ??= Array.Empty<int[]>();
        EvaluationCost = Math.Max(1, functions?.EvaluationCost ?? 1);
    }

    public int EvaluationCost { get; }

    public double[]? PaintBounds { get; }

    internal int TriangleCount => _triangles.Length / 3;

    public int GetEvaluationCost(double x, double y) {
        int column = (int)Math.Floor((x - _left) / _cellWidth), row = (int)Math.Floor((y - _top) / _cellHeight);
        if (column < 0 || row < 0 || column > _columns || row > _rows) return EvaluationCost;
        int candidates = _cells[Math.Min(row, _rows - 1) * _columns + Math.Min(column, _columns - 1)].Length;
        return (int)Math.Min(int.MaxValue, (long)EvaluationCost + candidates);
    }

    public bool TrySampleComponents(double x, double y, double[] input, double[] components) {
        int column = (int)Math.Floor((x - _left) / _cellWidth), row = (int)Math.Floor((y - _top) / _cellHeight);
        if (column < 0 || row < 0 || column > _columns || row > _rows) return false;
        column = Math.Min(column, _columns - 1); row = Math.Min(row, _rows - 1);
        int[] candidates = _cells[row * _columns + column];
        // Cells list triangles in descending stream order, so the first hit is the one painted last.
        int best = -1;
        double w0 = 0D, w1 = 0D, w2 = 0D;
        for (int index = 0; index < candidates.Length; index++) {
            if (!TryBarycentric(candidates[index], x, y, out w0, out w1, out w2)) continue;
            best = candidates[index];
            break;
        }
        if (best < 0) return false;
        int v0 = _triangles[best * 3], v1 = _triangles[best * 3 + 1], v2 = _triangles[best * 3 + 2];
        if (_functions != null) {
            double t = w0 * _values[v0] + w1 * _values[v1] + w2 * _values[v2];
            return _functions.TryEvaluate(t, input, components);
        }
        for (int component = 0; component < _valueCount && component < components.Length; component++) {
            components[component] = w0 * _values[v0 * _valueCount + component] + w1 * _values[v1 * _valueCount + component] +
                w2 * _values[v2 * _valueCount + component];
        }
        return true;
    }

    private bool TryBarycentric(int triangle, double x, double y, out double w0, out double w1, out double w2) {
        int a = _triangles[triangle * 3], b = _triangles[triangle * 3 + 1], c = _triangles[triangle * 3 + 2];
        double x0 = _x[a], y0 = _y[a], x1 = _x[b], y1 = _y[b], x2 = _x[c], y2 = _y[c];
        double denominator = (y1 - y2) * (x0 - x2) + (x2 - x1) * (y0 - y2);
        w0 = w1 = w2 = 0D;
        if (Math.Abs(denominator) < 1e-18D) return false;
        w0 = ((y1 - y2) * (x - x2) + (x2 - x1) * (y - y2)) / denominator;
        w1 = ((y2 - y0) * (x - x2) + (x0 - x2) * (y - y2)) / denominator;
        w2 = 1D - w0 - w1;
        const double tolerance = -1e-9D;
        return w0 >= tolerance && w1 >= tolerance && w2 >= tolerance;
    }

    private void CellRange(int triangle, out int c0, out int r0, out int c1, out int r1) {
        int a = _triangles[triangle * 3], b = _triangles[triangle * 3 + 1], c = _triangles[triangle * 3 + 2];
        double minX = Math.Min(_x[a], Math.Min(_x[b], _x[c])), maxX = Math.Max(_x[a], Math.Max(_x[b], _x[c]));
        double minY = Math.Min(_y[a], Math.Min(_y[b], _y[c])), maxY = Math.Max(_y[a], Math.Max(_y[b], _y[c]));
        c0 = Clamp((int)Math.Floor((minX - _left) / _cellWidth), _columns);
        c1 = Clamp((int)Math.Floor((maxX - _left) / _cellWidth), _columns);
        r0 = Clamp((int)Math.Floor((minY - _top) / _cellHeight), _rows);
        r1 = Clamp((int)Math.Floor((maxY - _top) / _cellHeight), _rows);
    }

    private static int Clamp(int value, int count) => value < 0 ? 0 : value >= count ? count - 1 : value;
}
