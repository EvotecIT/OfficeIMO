using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Interactive retained PDF page surface with text selection, copy, and link activation.</summary>
public sealed class PdfPageCanvas : Control, IDisposable {
    public static readonly StyledProperty<PdfPageScene?> SceneProperty =
        AvaloniaProperty.Register<PdfPageCanvas, PdfPageScene?>(nameof(Scene));

    public static readonly StyledProperty<Bitmap?> FallbackImageProperty =
        AvaloniaProperty.Register<PdfPageCanvas, Bitmap?>(nameof(FallbackImage));

    private readonly OfficeDrawingAvaloniaRenderer _renderer = new();
    private readonly Cursor _textCursor = new(StandardCursorType.Ibeam);
    private readonly Cursor _handCursor = new(StandardCursorType.Hand);
    private Point? _selectionStart;
    private Point? _selectionEnd;
    private PdfPageInteractionRegion? _hoverRegion;
    private bool _selecting;
    private bool _disposed;

    static PdfPageCanvas() {
        AffectsRender<PdfPageCanvas>(SceneProperty, FallbackImageProperty);
    }

    public PdfPageCanvas() {
        Focusable = true;
        Cursor = _textCursor;
    }

    public PdfPageScene? Scene {
        get => GetValue(SceneProperty);
        set => SetValue(SceneProperty, value);
    }

    public Bitmap? FallbackImage {
        get => GetValue(FallbackImageProperty);
        set => SetValue(FallbackImageProperty, value);
    }

    internal string SelectedText {
        get {
            if (Scene is null || !_selectionStart.HasValue || !_selectionEnd.HasValue) return string.Empty;
            Point start = ToPagePoint(_selectionStart.Value);
            Point end = ToPagePoint(_selectionEnd.Value);
            return Scene.Interactions.GetSelectedText(start.X, start.Y, end.X, end.Y);
        }
    }

    internal event Action<string>? LinkActivated;

    public override void Render(DrawingContext context) {
        base.Render(context);
        context.DrawRectangle(Brushes.White, null, Bounds);
        PdfPageScene? scene = Scene;
        if (scene is null) return;

        double scaleX = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double scaleY = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        using IDisposable scale = context.PushTransform(Matrix.CreateScale(scaleX, scaleY));
        if (scene.RequiresRasterFallback && FallbackImage is not null) {
            context.DrawImage(FallbackImage, new Rect(0, 0, scene.Drawing.Width, scene.Drawing.Height));
        } else if (!scene.RequiresRasterFallback) {
            _renderer.Render(context, scene.Drawing);
        }

        DrawSelection(context, scene);
        DrawInteractionOverlay(context);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _renderer.Dispose();
        _textCursor.Dispose();
        _handCursor.Dispose();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change) {
        base.OnPropertyChanged(change);
        if (change.Property != SceneProperty) return;
        _renderer.ClearImages();
        _selectionStart = null;
        _selectionEnd = null;
        _hoverRegion = null;
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e) {
        base.OnDetachedFromVisualTree(e);
        _renderer.ClearImages();
        _selecting = false;
        _selectionStart = null;
        _selectionEnd = null;
        _hoverRegion = null;
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e) {
        base.OnPointerPressed(e);
        if (Scene is null || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        Focus();
        _selectionStart = e.GetPosition(this);
        _selectionEnd = _selectionStart;
        _selecting = true;
        e.Pointer.Capture(this);
        e.Handled = true;
        InvalidateVisual();
    }

    protected override void OnPointerMoved(PointerEventArgs e) {
        base.OnPointerMoved(e);
        if (_selecting) {
            _selectionEnd = e.GetPosition(this);
            e.Handled = true;
            InvalidateVisual();
            return;
        }

        PdfPageInteractionRegion? previous = _hoverRegion;
        _hoverRegion = HitTestInteractive(e.GetPosition(this));
        if (!ReferenceEquals(previous, _hoverRegion)) {
            Cursor = _hoverRegion?.Kind == PdfInteractionKind.Link ? _handCursor : _textCursor;
            InvalidateVisual();
        }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e) {
        base.OnPointerReleased(e);
        if (!_selecting || !_selectionStart.HasValue) return;
        _selectionEnd = e.GetPosition(this);
        _selecting = false;
        e.Pointer.Capture(null);
        e.Handled = true;

        double horizontalDistance = _selectionEnd.Value.X - _selectionStart.Value.X;
        double verticalDistance = _selectionEnd.Value.Y - _selectionStart.Value.Y;
        if (Math.Sqrt((horizontalDistance * horizontalDistance) + (verticalDistance * verticalDistance)) < 4D) {
            ActivateLink(_selectionEnd.Value);
        }
        InvalidateVisual();
    }

    protected override void OnPointerExited(PointerEventArgs e) {
        base.OnPointerExited(e);
        if (_selecting || _hoverRegion is null) return;
        _hoverRegion = null;
        Cursor = _textCursor;
        InvalidateVisual();
    }

    protected override void OnKeyDown(KeyEventArgs e) {
        base.OnKeyDown(e);
        if (e.Key == Key.C &&
            (e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta)) &&
            !string.IsNullOrEmpty(SelectedText)) {
            _ = CopySelectionAsync();
            e.Handled = true;
        } else if (e.Key == Key.Escape) {
            _selectionStart = null;
            _selectionEnd = null;
            InvalidateVisual();
            e.Handled = true;
        }
    }

    private void ActivateLink(Point controlPoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return;
        Point point = ToPagePoint(controlPoint);
        PdfPageInteractionRegion? link = scene.Interactions
            .HitTest(point.X, point.Y, tolerance: 1D)
            .FirstOrDefault(static region => region.Kind == PdfInteractionKind.Link);
        if (!string.IsNullOrWhiteSpace(link?.Target)) LinkActivated?.Invoke(link.Target!);
    }

    private PdfPageInteractionRegion? HitTestInteractive(Point controlPoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return null;
        Point point = ToPagePoint(controlPoint);
        return scene.Interactions.HitTest(point.X, point.Y, tolerance: 1D)
            .FirstOrDefault(static region => region.Kind != PdfInteractionKind.Text);
    }

    private async Task CopySelectionAsync() {
        IClipboard? clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is not null) await clipboard.SetTextAsync(SelectedText);
    }

    private void DrawSelection(DrawingContext context, PdfPageScene scene) {
        if (!_selectionStart.HasValue || !_selectionEnd.HasValue) return;
        Point start = ToPagePoint(_selectionStart.Value);
        Point end = ToPagePoint(_selectionEnd.Value);
        var brush = new SolidColorBrush(Color.FromArgb(72, 53, 106, 230));
        foreach (PdfPageInteractionRegion region in scene.Interactions.SelectText(start.X, start.Y, end.X, end.Y)) {
            context.DrawRectangle(
                brush,
                null,
                new Rect(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height));
        }
    }

    private void DrawInteractionOverlay(DrawingContext context) {
        if (_hoverRegion is null) return;
        Color color = _hoverRegion.Kind switch {
            PdfInteractionKind.Link => Color.FromRgb(53, 106, 230),
            PdfInteractionKind.FormWidget => Color.FromRgb(16, 185, 129),
            _ => Color.FromRgb(245, 158, 11)
        };
        var fill = new SolidColorBrush(Color.FromArgb(28, color.R, color.G, color.B));
        var stroke = new Pen(new SolidColorBrush(Color.FromArgb(190, color.R, color.G, color.B)), 1D);
        context.DrawRectangle(
            fill,
            stroke,
            new Rect(_hoverRegion.Quad.Left, _hoverRegion.Quad.Top, _hoverRegion.Quad.Width, _hoverRegion.Quad.Height));
    }

    private Point ToPagePoint(Point controlPoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return default;
        double scaleX = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double scaleY = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        return new Point(controlPoint.X / Math.Max(scaleX, 0.000001D), controlPoint.Y / Math.Max(scaleY, 0.000001D));
    }
}
