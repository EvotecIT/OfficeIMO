using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Draws an OCR region or adjusts four perspective corners in source-image coordinates.</summary>
public sealed class ScanSelectionCanvas : Control {
    private ScanPreparationViewModel? _model;
    private Point? _start;
    private int _corner = -1;
    private int _keyboardCorner;
    public ScanSelectionCanvas() {
        Focusable = true;
        ClipToBounds = true;
        DataContextChanged += (_, _) => {
            if (_model != null) _model.PropertyChanged -= ModelChanged;
            _model = DataContext as ScanPreparationViewModel;
            if (_model != null) _model.PropertyChanged += ModelChanged;
            InvalidateVisual();
        };
        DetachedFromVisualTree += (_, _) => { if (_model != null) _model.PropertyChanged -= ModelChanged; };
        AttachedToVisualTree += (_, _) => { if (_model != null) { _model.PropertyChanged -= ModelChanged; _model.PropertyChanged += ModelChanged; } };
    }
    private void ModelChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName is nameof(ScanPreparationViewModel.SourcePreview)
            or nameof(ScanPreparationViewModel.EditCorners) or nameof(ScanPreparationViewModel.UsePerspective)) {
            _start = null;
            _corner = -1;
        }
        InvalidateVisual();
    }
    private Rect ImageBounds() {
        if (_model?.SourcePreview is not { } image) return default;
        double scale = Math.Min(Bounds.Width / image.Size.Width, Bounds.Height / image.Size.Height);
        var size = new Size(image.Size.Width * scale, image.Size.Height * scale);
        return new Rect((Bounds.Width - size.Width) / 2, (Bounds.Height - size.Height) / 2, size.Width, size.Height);
    }
    public override void Render(DrawingContext context) {
        base.Render(context);
        if (_model?.SourcePreview is not { } image) return;
        Rect bounds = ImageBounds();
        context.DrawImage(image, new Rect(image.Size), bounds);
        context.DrawRectangle(null, new Pen(Brushes.Gray, 1), bounds);
        Rect region = _model.UseRegion ? _model.Region : new Rect(0, 0, 1, 1);
        Rect rectangle = new(bounds.X + region.X * bounds.Width, bounds.Y + region.Y * bounds.Height, region.Width * bounds.Width, region.Height * bounds.Height);
        var pen = new Pen(Brushes.DodgerBlue, 2);
        if (_model.UseRegion) context.DrawRectangle(new SolidColorBrush(Color.FromArgb(28, 0, 120, 255)), pen, rectangle);
        if (_model.UsePerspective) {
            Point[] corners = { _model.TopLeft, _model.TopRight, _model.BottomRight, _model.BottomLeft };
            Point Map(Point p) => new(rectangle.X + p.X * rectangle.Width, rectangle.Y + p.Y * rectangle.Height);
            for (int i = 0; i < 4; i++) {
                context.DrawLine(pen, Map(corners[i]), Map(corners[(i + 1) % 4]));
                context.DrawEllipse(IsFocused && _model.EditCorners && i == _keyboardCorner ? Brushes.DodgerBlue : Brushes.White, pen, Map(corners[i]), 6, 6);
            }
        }
    }
    protected override void OnPointerPressed(PointerPressedEventArgs e) {
        base.OnPointerPressed(e);
        if (_model?.CanSelectRegion != true || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        Rect bounds = ImageBounds(); Point position = e.GetPosition(this);
        if (!bounds.Contains(position)) return;
        Focus();
        _start = Normalize(position, bounds);
        if (_model.UsePerspective && _model.EditCorners) {
            Point normalized = InRegion(_start.Value);
            Point[] corners = { _model.TopLeft, _model.TopRight, _model.BottomRight, _model.BottomLeft };
            _corner = Enumerable.Range(0, 4).OrderBy(i => Math.Pow(corners[i].X - normalized.X, 2) + Math.Pow(corners[i].Y - normalized.Y, 2)).First();
            _keyboardCorner = _corner;
            SetCorner(normalized);
        } else { _corner = -1; _model.UseRegion = true; }
        e.Pointer.Capture(this); e.Handled = true;
    }
    protected override void OnPointerMoved(PointerEventArgs e) {
        base.OnPointerMoved(e);
        if (_start == null || _model?.CanSelectRegion != true) return;
        Point point = Normalize(e.GetPosition(this), ImageBounds());
        if (_corner >= 0) SetCorner(InRegion(point));
        else {
            double left = Math.Min(_start.Value.X, point.X), top = Math.Min(_start.Value.Y, point.Y);
            double width = Math.Abs(point.X - _start.Value.X), height = Math.Abs(point.Y - _start.Value.Y);
            if (width > .002 && height > .002) _model.Region = new Rect(left, top, width, height);
        }
        e.Handled = true;
    }
    protected override void OnPointerReleased(PointerReleasedEventArgs e) {
        base.OnPointerReleased(e); if (_start == null) return;
        _start = null; _corner = -1; e.Pointer.Capture(null); e.Handled = true;
    }
    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e) { _start = null; _corner = -1; base.OnPointerCaptureLost(e); }
    protected override void OnKeyDown(KeyEventArgs e) {
        base.OnKeyDown(e);
        if (_model?.CanSelectRegion != true) return;
        if (_model.UsePerspective && _model.EditCorners && e.Key is >= Key.D1 and <= Key.D4) {
            _keyboardCorner = (int)e.Key - (int)Key.D1;
            e.Handled = true;
            InvalidateVisual();
            return;
        }
        double dx = e.Key == Key.Left ? -0.01 : e.Key == Key.Right ? 0.01 : 0;
        double dy = e.Key == Key.Up ? -0.01 : e.Key == Key.Down ? 0.01 : 0;
        if (dx == 0 && dy == 0) return;
        if (_model.UsePerspective && _model.EditCorners) {
            Point[] corners = { _model.TopLeft, _model.TopRight, _model.BottomRight, _model.BottomLeft };
            Point current = corners[_keyboardCorner];
            _corner = _keyboardCorner;
            SetCorner(new Point(Math.Clamp(current.X + dx, 0, 1), Math.Clamp(current.Y + dy, 0, 1)));
            _corner = -1;
        } else {
            _model.UseRegion = true;
            Rect current = _model.Region;
            _model.Region = e.KeyModifiers.HasFlag(KeyModifiers.Shift)
                ? new Rect(current.X, current.Y, Math.Clamp(current.Width + dx, Math.Min(0.01, 1 - current.X), 1 - current.X),
                    Math.Clamp(current.Height + dy, Math.Min(0.01, 1 - current.Y), 1 - current.Y))
                : new Rect(Math.Clamp(current.X + dx, 0, 1 - current.Width),
                    Math.Clamp(current.Y + dy, 0, 1 - current.Height), current.Width, current.Height);
        }
        e.Handled = true;
    }
    private static Point Normalize(Point point, Rect bounds) => new(Math.Clamp((point.X - bounds.X) / bounds.Width, 0, 1), Math.Clamp((point.Y - bounds.Y) / bounds.Height, 0, 1));
    private Point InRegion(Point point) {
        Rect region = _model!.UseRegion ? _model.Region : new Rect(0, 0, 1, 1);
        return new(Math.Clamp((point.X - region.X) / region.Width, 0, 1), Math.Clamp((point.Y - region.Y) / region.Height, 0, 1));
    }
    private void SetCorner(Point point) {
        switch (_corner) { case 0: _model!.TopLeft = point; break; case 1: _model!.TopRight = point; break; case 2: _model!.BottomRight = point; break; case 3: _model!.BottomLeft = point; break; }
    }
}
