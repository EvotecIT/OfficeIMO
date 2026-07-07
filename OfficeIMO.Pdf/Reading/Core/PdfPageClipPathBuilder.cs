using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageClipPathBuilder {
    private readonly double _pageHeight;
    private readonly List<(double X, double Y)> _path = new List<(double X, double Y)>();
    private readonly List<OfficePathCommand> _pathCommands = new List<OfficePathCommand>();
    private int _currentSubpathStartIndex = -1;
    private bool _currentSubpathHasDraw;

    public PdfPageClipPathBuilder(double pageHeight) {
        _pageHeight = pageHeight;
    }

    public void AddRectanglePath(Matrix2D transform, double x, double y, double width, double height) {
        DiscardCurrentSubpathIfEmpty();
        var p0 = transform.Transform(x, y);
        var p1 = transform.Transform(x + width, y);
        var p2 = transform.Transform(x + width, y + height);
        var p3 = transform.Transform(x, y + height);
        _currentSubpathStartIndex = _path.Count;
        _currentSubpathHasDraw = true;
        _path.Add(p0);
        _path.Add(p1);
        _path.Add(p2);
        _path.Add(p3);
        _path.Add(p0);
        _pathCommands.Add(OfficePathCommand.MoveTo(ToOfficePoint(p0)));
        _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p1)));
        _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p2)));
        _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(p3)));
        _pathCommands.Add(OfficePathCommand.Close());
    }

    public void MoveTo(Matrix2D transform, double x, double y) {
        DiscardCurrentSubpathIfEmpty();
        (double X, double Y) point = transform.Transform(x, y);
        _currentSubpathStartIndex = _path.Count;
        _currentSubpathHasDraw = false;
        _path.Add(point);
        _pathCommands.Add(OfficePathCommand.MoveTo(ToOfficePoint(point)));
    }

    public void LineTo(Matrix2D transform, double x, double y) {
        if (_currentSubpathStartIndex < 0) {
            MoveTo(transform, x, y);
            return;
        }

        (double X, double Y) point = transform.Transform(x, y);
        _path.Add(point);
        _currentSubpathHasDraw = true;
        _pathCommands.Add(OfficePathCommand.LineTo(ToOfficePoint(point)));
    }

    public void CubicTo(Matrix2D transform, double c1x, double c1y, double c2x, double c2y, double endX, double endY) {
        if (_path.Count == 0 || _currentSubpathStartIndex < 0) {
            MoveTo(transform, endX, endY);
            return;
        }

        (double X, double Y) control1 = transform.Transform(c1x, c1y);
        (double X, double Y) control2 = transform.Transform(c2x, c2y);
        (double X, double Y) end = transform.Transform(endX, endY);
        AddCubic(control1, control2, end);
    }

    public void CubicToWithCurrentFirstControl(Matrix2D transform, double c2x, double c2y, double endX, double endY) {
        if (_path.Count == 0 || _currentSubpathStartIndex < 0) {
            MoveTo(transform, endX, endY);
            return;
        }

        (double X, double Y) control1 = _path[_path.Count - 1];
        (double X, double Y) control2 = transform.Transform(c2x, c2y);
        (double X, double Y) end = transform.Transform(endX, endY);
        AddCubic(control1, control2, end);
    }

    public void CubicToWithEndSecondControl(Matrix2D transform, double c1x, double c1y, double endX, double endY) {
        if (_path.Count == 0 || _currentSubpathStartIndex < 0) {
            MoveTo(transform, endX, endY);
            return;
        }

        (double X, double Y) control1 = transform.Transform(c1x, c1y);
        (double X, double Y) end = transform.Transform(endX, endY);
        AddCubic(control1, end, end);
    }

    public void ClosePath() {
        if (_path.Count == 0 || _currentSubpathStartIndex < 0 || _currentSubpathStartIndex >= _path.Count || !_currentSubpathHasDraw) {
            return;
        }

        _path.Add(_path[_currentSubpathStartIndex]);
        _pathCommands.Add(OfficePathCommand.Close());
    }

    public void Clear() {
        _path.Clear();
        _pathCommands.Clear();
        _currentSubpathStartIndex = -1;
        _currentSubpathHasDraw = false;
    }

    public bool TryCreateClipPath(OfficeFillRule fillRule, out PdfPageClipPath clipPath) {
        clipPath = default;
        if (_path.Count < 2) {
            return false;
        }

        if (TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height)) {
            clipPath = PdfPageClipPath.Rectangle(x, y, width, height);
            return true;
        }

        return PdfPageClipPath.TryCreatePath(_pathCommands, fillRule, out clipPath);
    }

    private void AddCubic((double X, double Y) control1, (double X, double Y) control2, (double X, double Y) end) {
        _path.Add(end);
        _currentSubpathHasDraw = true;
        _pathCommands.Add(OfficePathCommand.CubicBezierTo(ToOfficePoint(control1), ToOfficePoint(control2), ToOfficePoint(end)));
    }

    private bool TryCreateAxisAlignedRectangle(out double x, out double y, out double width, out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (_path.Count != 5 ||
            _pathCommands.Count != 5 ||
            CountMoveCommands() != 1 ||
            _pathCommands[0].Kind != OfficePathCommandKind.MoveTo ||
            _pathCommands[1].Kind != OfficePathCommandKind.LineTo ||
            _pathCommands[2].Kind != OfficePathCommandKind.LineTo ||
            _pathCommands[3].Kind != OfficePathCommandKind.LineTo ||
            _pathCommands[4].Kind != OfficePathCommandKind.Close ||
            !NearlyEqual(_path[0].X, _path[4].X) ||
            !NearlyEqual(_path[0].Y, _path[4].Y)) {
            return false;
        }

        double left = _path.Min(point => point.X);
        double right = _path.Max(point => point.X);
        double top = _path.Min(point => ToTop(point.Y));
        double bottom = _path.Max(point => ToTop(point.Y));
        width = right - left;
        height = bottom - top;
        if (width <= 0D || height <= 0D) {
            return false;
        }

        var corners = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < 4; i++) {
            bool onVertical = NearlyEqual(_path[i].X, left) || NearlyEqual(_path[i].X, right);
            bool onHorizontal = NearlyEqual(ToTop(_path[i].Y), top) || NearlyEqual(ToTop(_path[i].Y), bottom);
            if (!onVertical || !onHorizontal) {
                return false;
            }

            (double X, double Y) next = _path[i + 1];
            bool horizontalEdge = NearlyEqual(_path[i].Y, next.Y) && !NearlyEqual(_path[i].X, next.X);
            bool verticalEdge = NearlyEqual(_path[i].X, next.X) && !NearlyEqual(_path[i].Y, next.Y);
            if (!horizontalEdge && !verticalEdge) {
                return false;
            }

            corners.Add((NearlyEqual(_path[i].X, left) ? "L" : "R") + (NearlyEqual(ToTop(_path[i].Y), top) ? "T" : "B"));
        }

        if (corners.Count != 4) {
            return false;
        }

        x = left;
        y = top;
        return true;
    }

    private void DiscardCurrentSubpathIfEmpty() {
        if (_currentSubpathHasDraw ||
            _currentSubpathStartIndex < 0 ||
            _currentSubpathStartIndex >= _path.Count) {
            return;
        }

        _path.RemoveRange(_currentSubpathStartIndex, _path.Count - _currentSubpathStartIndex);
        if (_pathCommands.Count > 0 && _pathCommands[_pathCommands.Count - 1].Kind == OfficePathCommandKind.MoveTo) {
            _pathCommands.RemoveAt(_pathCommands.Count - 1);
        }

        _currentSubpathStartIndex = -1;
    }

    private int CountMoveCommands() {
        int count = 0;
        for (int i = 0; i < _pathCommands.Count; i++) {
            if (_pathCommands[i].Kind == OfficePathCommandKind.MoveTo) {
                count++;
            }
        }

        return count;
    }

    private OfficePoint ToOfficePoint((double X, double Y) point) =>
        new OfficePoint(point.X, ToTop(point.Y));

    private double ToTop(double y) => _pageHeight - y;

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
}
