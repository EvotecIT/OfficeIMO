using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;

namespace OfficeIMO.Studio.Features.Sign;

/// <summary>Captures a handwritten signature as strokes and renders it to a transparent PNG.</summary>
public sealed class SignaturePad : Control {
    internal static readonly Color InkColor = Color.FromRgb(27, 42, 74);
    private const double StrokeThickness = 2.6D;
    private readonly List<List<Point>> _strokes = [];
    private List<Point>? _current;

    public SignaturePad() {
        Cursor = new Cursor(StandardCursorType.Cross);
        Focusable = true;
    }

    public event EventHandler? InkChanged;

    public bool HasInk => _strokes.Any(stroke => stroke.Count > 1);

    public void Clear() {
        _strokes.Clear();
        _current = null;
        InvalidateVisual();
        InkChanged?.Invoke(this, EventArgs.Empty);
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e) {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        _current = [e.GetPosition(this)];
        _strokes.Add(_current);
        e.Pointer.Capture(this);
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e) {
        base.OnPointerMoved(e);
        if (_current is null) return;
        Point point = e.GetPosition(this);
        if (_current.Count > 0 && Distance(_current[^1], point) < 1.2D) return;
        _current.Add(point);
        InvalidateVisual();
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e) {
        base.OnPointerReleased(e);
        if (_current is null) return;
        if (_current.Count == 1) _current.Add(_current[0] + new Point(0.8D, 0.8D));
        _current = null;
        e.Pointer.Capture(null);
        InvalidateVisual();
        InkChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Adds a stroke programmatically, for automation and tests.</summary>
    internal void AddStroke(IEnumerable<Point> points) {
        _strokes.Add(points.ToList());
        InvalidateVisual();
        InkChanged?.Invoke(this, EventArgs.Empty);
    }

    public override void Render(DrawingContext context) {
        base.Render(context);
        context.FillRectangle(Brushes.Transparent, new Rect(Bounds.Size));
        DrawStrokes(context, new Pen(new SolidColorBrush(InkColor), StrokeThickness, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
    }

    private void DrawStrokes(DrawingContext context, IPen pen) {
        foreach (List<Point> stroke in _strokes.Where(stroke => stroke.Count > 1)) {
            var geometry = new StreamGeometry();
            using (StreamGeometryContext path = geometry.Open()) {
                path.BeginFigure(stroke[0], isFilled: false);
                for (int i = 1; i < stroke.Count; i++) path.LineTo(stroke[i]);
                path.EndFigure(isClosed: false);
            }
            context.DrawGeometry(null, pen, geometry);
        }
    }

    /// <summary>Strokes in the coordinate space of <see cref="ToPng"/>'s crop, scaled to 0..1 on both axes.</summary>
    internal IReadOnlyList<IReadOnlyList<Point>>? NormalizedStrokes() {
        Point[] points = _strokes.Where(stroke => stroke.Count > 1).SelectMany(stroke => stroke).ToArray();
        if (points.Length == 0) return null;
        const double padding = 6D;
        double left = points.Min(point => point.X) - padding, top = points.Min(point => point.Y) - padding;
        double width = points.Max(point => point.X) - left + padding, height = points.Max(point => point.Y) - top + padding;
        return _strokes.Where(stroke => stroke.Count > 1)
            .Select(stroke => (IReadOnlyList<Point>)stroke.Select(point => new Point((point.X - left) / width, (point.Y - top) / height)).ToArray())
            .ToArray();
    }

    /// <summary>Renders the ink, cropped to its bounds, at three times screen resolution.</summary>
    internal byte[]? ToPng() {
        Point[] points = _strokes.Where(stroke => stroke.Count > 1).SelectMany(stroke => stroke).ToArray();
        if (points.Length == 0) return null;
        const double padding = 6D, scale = 3D;
        double left = points.Min(point => point.X) - padding, top = points.Min(point => point.Y) - padding;
        double width = points.Max(point => point.X) - left + padding, height = points.Max(point => point.Y) - top + padding;
        var size = new PixelSize(Math.Max(1, (int)Math.Ceiling(width * scale)), Math.Max(1, (int)Math.Ceiling(height * scale)));
        using var bitmap = new RenderTargetBitmap(size, new Vector(96D, 96D));
        using (DrawingContext context = bitmap.CreateDrawingContext()) {
            using (context.PushTransform(Matrix.CreateTranslation(-left, -top) * Matrix.CreateScale(scale, scale))) {
                DrawStrokes(context, new Pen(new SolidColorBrush(InkColor), StrokeThickness, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
            }
        }
        using var stream = new MemoryStream();
        bitmap.Save(stream, PngBitmapEncoderOptions.Default);
        return stream.ToArray();
    }

    private static double Distance(Point a, Point b) => Math.Sqrt(Math.Pow(a.X - b.X, 2) + Math.Pow(a.Y - b.Y, 2));
}
