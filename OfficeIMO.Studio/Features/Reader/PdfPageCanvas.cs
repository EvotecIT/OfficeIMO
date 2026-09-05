using Avalonia;
using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Interactive retained PDF page surface with text selection, copy, and link activation.</summary>
public sealed partial class PdfPageCanvas : Control, IDisposable {
    public static readonly StyledProperty<PdfPageScene?> SceneProperty =
        AvaloniaProperty.Register<PdfPageCanvas, PdfPageScene?>(nameof(Scene));

    public static readonly StyledProperty<Bitmap?> FallbackImageProperty =
        AvaloniaProperty.Register<PdfPageCanvas, Bitmap?>(nameof(FallbackImage));

    public static readonly StyledProperty<PdfEditorTool> EditorToolProperty =
        AvaloniaProperty.Register<PdfPageCanvas, PdfEditorTool>(nameof(EditorTool), PdfEditorTool.Select);

    public static readonly StyledProperty<PdfEditorSelection?> SelectedObjectProperty =
        AvaloniaProperty.Register<PdfPageCanvas, PdfEditorSelection?>(nameof(SelectedObject));

    public static readonly StyledProperty<PdfEditorSelectionMode> SelectionModeProperty =
        AvaloniaProperty.Register<PdfPageCanvas, PdfEditorSelectionMode>(nameof(SelectionMode));

    public static readonly StyledProperty<Rect?> PendingRedactionAreaProperty =
        AvaloniaProperty.Register<PdfPageCanvas, Rect?>(nameof(PendingRedactionArea));

    private readonly OfficeDrawingAvaloniaRenderer _renderer = new();
    private readonly Cursor _textCursor = new(StandardCursorType.Ibeam);
    private readonly Cursor _handCursor = new(StandardCursorType.Hand);
    private readonly Cursor _crossCursor = new(StandardCursorType.Cross);
    private readonly List<Point> _editorPath = new();
    private Point? _selectionStart;
    private Point? _selectionEnd;
    private PdfPageInteractionRegion? _hoverRegion;
    private bool _selecting;
    private bool _editing;
    private bool _disposed;
    private int _keyboardInteractionIndex = -1;
    private PdfPageCanvasAutomationPeer? _automationPeer;

    static PdfPageCanvas() {
        AffectsRender<PdfPageCanvas>(
            SceneProperty,
            FallbackImageProperty,
            EditorToolProperty,
            SelectedObjectProperty,
            SelectionModeProperty,
            PendingRedactionAreaProperty);
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

    public PdfEditorTool EditorTool {
        get => GetValue(EditorToolProperty);
        set => SetValue(EditorToolProperty, value);
    }

    public PdfEditorSelection? SelectedObject {
        get => GetValue(SelectedObjectProperty);
        set => SetValue(SelectedObjectProperty, value);
    }

    public PdfEditorSelectionMode SelectionMode {
        get => GetValue(SelectionModeProperty);
        set => SetValue(SelectionModeProperty, value);
    }

    public Rect? PendingRedactionArea {
        get => GetValue(PendingRedactionAreaProperty);
        set => SetValue(PendingRedactionAreaProperty, value);
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

    internal event Action<PdfEditorGesture>? EditorGestureCompleted;

    internal event Action<PdfEditorSelection?>? ObjectSelected;

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
        DrawSelectedObject(context);
        DrawPendingRedaction(context);
        DrawEditorPreview(context);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _renderer.Dispose();
        _textCursor.Dispose();
        _handCursor.Dispose();
        _crossCursor.Dispose();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change) {
        base.OnPropertyChanged(change);
        if (change.Property == EditorToolProperty) {
            Cursor = EditorTool == PdfEditorTool.Select ? _textCursor : _crossCursor;
            ResetPointerState();
            InvalidateVisual();
            return;
        }
        if (change.Property != SceneProperty) return;
        _renderer.ClearImages();
        _selectionStart = null;
        _selectionEnd = null;
        _hoverRegion = null;
        _editorPath.Clear();
        _keyboardInteractionIndex = -1;
        _automationPeer?.RefreshChildren();
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e) {
        base.OnDetachedFromVisualTree(e);
        _renderer.ClearImages();
        _selecting = false;
        _editing = false;
        _selectionStart = null;
        _selectionEnd = null;
        _hoverRegion = null;
        _editorPath.Clear();
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e) {
        base.OnPointerPressed(e);
        if (Scene is null || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        Focus();
        if (EditorTool != PdfEditorTool.Select) {
            _editorPath.Clear();
            _editorPath.Add(ToPagePoint(e.GetPosition(this)));
            _editing = true;
            e.Pointer.Capture(this);
            e.Handled = true;
            InvalidateVisual();
            return;
        }
        _selectionStart = e.GetPosition(this);
        _selectionEnd = _selectionStart;
        _selecting = true;
        e.Pointer.Capture(this);
        e.Handled = true;
        InvalidateVisual();
    }

    protected override void OnPointerMoved(PointerEventArgs e) {
        base.OnPointerMoved(e);
        if (_editing) {
            Point point = ToPagePoint(e.GetPosition(this));
            if (_editorPath.Count == 0 || Distance(_editorPath[^1], point) >= 1D) _editorPath.Add(point);
            e.Handled = true;
            InvalidateVisual();
            return;
        }
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
        if (_editing) {
            Point point = ToPagePoint(e.GetPosition(this));
            if (_editorPath.Count == 0 || Distance(_editorPath[^1], point) > 0.1D) _editorPath.Add(point);
            _editing = false;
            e.Pointer.Capture(null);
            e.Handled = true;
            EmitEditorGesture();
            _editorPath.Clear();
            InvalidateVisual();
            return;
        }
        if (!_selecting || !_selectionStart.HasValue) return;
        _selectionEnd = e.GetPosition(this);
        _selecting = false;
        e.Pointer.Capture(null);
        e.Handled = true;

        double horizontalDistance = _selectionEnd.Value.X - _selectionStart.Value.X;
        double verticalDistance = _selectionEnd.Value.Y - _selectionStart.Value.Y;
        if (Math.Sqrt((horizontalDistance * horizontalDistance) + (verticalDistance * verticalDistance)) < 4D) {
            if (!SelectObjectAt(_selectionEnd.Value)) ActivateLink(_selectionEnd.Value);
        } else if (SelectionMode == PdfEditorSelectionMode.PageContent) {
            SelectTextObject();
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
            _editing = false;
            _editorPath.Clear();
            _selectionStart = null;
            _selectionEnd = null;
            ObjectSelected?.Invoke(null);
            InvalidateVisual();
            e.Handled = true;
        } else if (e.Key == Key.A &&
                   (e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta))) {
            SelectAllText();
            e.Handled = true;
        } else if (e.Key is Key.Left or Key.Up) {
            MoveKeyboardInteraction(-1);
            e.Handled = true;
        } else if (e.Key is Key.Right or Key.Down) {
            MoveKeyboardInteraction(1);
            e.Handled = true;
        } else if (e.Key is Key.Home) {
            MoveKeyboardInteractionTo(0);
            e.Handled = true;
        } else if (e.Key is Key.End) {
            MoveKeyboardInteractionTo(GetKeyboardInteractions().Count - 1);
            e.Handled = true;
        } else if (e.Key is Key.Enter or Key.Space) {
            ActivateKeyboardInteraction();
            e.Handled = true;
        }
    }

    protected override AutomationPeer OnCreateAutomationPeer() =>
        _automationPeer ??= new PdfPageCanvasAutomationPeer(this);

    private IReadOnlyList<PdfPageInteractionRegion> GetKeyboardInteractions() =>
        Scene?.Interactions.Regions.Where(static region => region.Kind != PdfInteractionKind.Text).ToArray()
        ?? Array.Empty<PdfPageInteractionRegion>();

    private void MoveKeyboardInteraction(int offset) {
        IReadOnlyList<PdfPageInteractionRegion> interactions = GetKeyboardInteractions();
        if (interactions.Count == 0) return;
        int next = _keyboardInteractionIndex < 0
            ? offset < 0 ? interactions.Count - 1 : 0
            : (_keyboardInteractionIndex + offset + interactions.Count) % interactions.Count;
        MoveKeyboardInteractionTo(next);
    }

    private void MoveKeyboardInteractionTo(int index) {
        IReadOnlyList<PdfPageInteractionRegion> interactions = GetKeyboardInteractions();
        if (index < 0 || index >= interactions.Count) return;
        _keyboardInteractionIndex = index;
        _hoverRegion = interactions[index];
        SelectRegion(interactions[index], activateLink: false);
        InvalidateVisual();
    }

    private void ActivateKeyboardInteraction() {
        IReadOnlyList<PdfPageInteractionRegion> interactions = GetKeyboardInteractions();
        if (_keyboardInteractionIndex < 0 || _keyboardInteractionIndex >= interactions.Count) return;
        SelectRegion(interactions[_keyboardInteractionIndex], activateLink: true);
    }

    internal void SelectRegion(PdfPageInteractionRegion region, bool activateLink) {
        if (region.Kind == PdfInteractionKind.Link) {
            if (activateLink && !string.IsNullOrWhiteSpace(region.Target)) LinkActivated?.Invoke(region.Target!);
            return;
        }
        if (Scene is not null) ObjectSelected?.Invoke(CreateSelection(Scene.PageNumber, region));
    }

    private void SelectAllText() {
        IReadOnlyList<PdfPageInteractionRegion>? regions = Scene?.Interactions.TextRegions;
        if (regions is null || regions.Count == 0) return;
        double left = regions.Min(static region => region.Quad.Left);
        double top = regions.Min(static region => region.Quad.Top);
        double right = regions.Max(static region => region.Quad.Right);
        double bottom = regions.Max(static region => region.Quad.Bottom);
        _selectionStart = ToControlPoint(new Point(left, top));
        _selectionEnd = ToControlPoint(new Point(right, bottom));
        SelectTextObject();
        InvalidateVisual();
    }

    private Point ToControlPoint(Point pagePoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return pagePoint;
        return new Point(
            pagePoint.X * Bounds.Width / Math.Max(1D, scene.Drawing.Width),
            pagePoint.Y * Bounds.Height / Math.Max(1D, scene.Drawing.Height));
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

    private bool SelectObjectAt(Point controlPoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return false;
        Point point = ToPagePoint(controlPoint);
        IReadOnlyList<PdfPageInteractionRegion> matches = scene.Interactions.HitTest(point.X, point.Y, tolerance: 2D);
        PdfPageInteractionRegion? selected = SelectionMode switch {
            PdfEditorSelectionMode.Annotations => matches.FirstOrDefault(static region =>
                region.Kind == PdfInteractionKind.Annotation && region.ObjectNumber.HasValue),
            PdfEditorSelectionMode.PageContent => matches.FirstOrDefault(static region =>
                (region.Kind == PdfInteractionKind.Annotation && region.ObjectNumber.HasValue) ||
                (region.Kind == PdfInteractionKind.Image && region.ImagePlacement is not null) ||
                region.Kind == PdfInteractionKind.Text),
            _ => null
        };
        if (selected is null) {
            ObjectSelected?.Invoke(null);
            return false;
        }

        ObjectSelected?.Invoke(CreateSelection(scene.PageNumber, selected));
        return true;
    }

    private void SelectTextObject() {
        PdfPageScene? scene = Scene;
        if (scene is null || !_selectionStart.HasValue || !_selectionEnd.HasValue) return;
        Point start = ToPagePoint(_selectionStart.Value);
        Point end = ToPagePoint(_selectionEnd.Value);
        IReadOnlyList<PdfPageInteractionRegion> regions = scene.Interactions.SelectText(start.X, start.Y, end.X, end.Y);
        if (regions.Count == 0) {
            ObjectSelected?.Invoke(null);
            return;
        }

        double left = regions.Min(static region => region.Quad.Left);
        double top = regions.Min(static region => region.Quad.Top);
        double right = regions.Max(static region => region.Quad.Right);
        double bottom = regions.Max(static region => region.Quad.Bottom);
        ObjectSelected?.Invoke(new PdfEditorSelection(
            PdfEditorSelectionKind.Text,
            scene.PageNumber,
            new PdfEditorVisualBounds(left, top, right, bottom),
            Text: string.Concat(regions.Select(static region => region.Text))));
    }

    private static PdfEditorSelection CreateSelection(int pageNumber, PdfPageInteractionRegion region) => new(
        region.Kind switch {
            PdfInteractionKind.Image => PdfEditorSelectionKind.Image,
            PdfInteractionKind.Annotation => PdfEditorSelectionKind.Annotation,
            _ => PdfEditorSelectionKind.Text
        },
        pageNumber,
        new PdfEditorVisualBounds(region.Quad.Left, region.Quad.Top, region.Quad.Right, region.Quad.Bottom),
        Text: region.Text,
        ObjectNumber: region.ObjectNumber,
        Subtype: region.Subtype,
        ImagePlacement: region.ImagePlacement);

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
            PdfInteractionKind.Image => Color.FromRgb(124, 58, 237),
            _ => Color.FromRgb(245, 158, 11)
        };
        var fill = new SolidColorBrush(Color.FromArgb(28, color.R, color.G, color.B));
        var stroke = new Pen(new SolidColorBrush(Color.FromArgb(190, color.R, color.G, color.B)), 1D);
        context.DrawRectangle(
            fill,
            stroke,
            new Rect(_hoverRegion.Quad.Left, _hoverRegion.Quad.Top, _hoverRegion.Quad.Width, _hoverRegion.Quad.Height));
    }

    private void DrawSelectedObject(DrawingContext context) {
        if (SelectedObject is not PdfEditorSelection selected) return;
        PdfEditorVisualBounds bounds = selected.Bounds;
        if (bounds.Width <= 0D || bounds.Height <= 0D) return;
        Color accent = selected.Kind switch {
            PdfEditorSelectionKind.Image => Color.FromRgb(124, 58, 237),
            PdfEditorSelectionKind.Annotation => Color.FromRgb(245, 158, 11),
            _ => Color.FromRgb(53, 106, 230)
        };
        var fill = new SolidColorBrush(Color.FromArgb(32, 53, 106, 230));
        var stroke = new Pen(new SolidColorBrush(accent), 2D);
        var area = new Rect(bounds.Left, bounds.Top, bounds.Width, bounds.Height);
        context.DrawRectangle(fill, stroke, area);
        DrawSelectionHandles(context, area, accent);
    }

    private static void DrawSelectionHandles(DrawingContext context, Rect area, Color accent) {
        const double handleSize = 7D;
        double half = handleSize / 2D;
        var fill = new SolidColorBrush(Colors.White);
        var stroke = new Pen(new SolidColorBrush(accent), 1.5D);
        Point[] handles = {
            area.TopLeft,
            new(area.Center.X, area.Top),
            area.TopRight,
            new(area.Right, area.Center.Y),
            area.BottomRight,
            new(area.Center.X, area.Bottom),
            area.BottomLeft,
            new(area.Left, area.Center.Y)
        };
        foreach (Point handle in handles) {
            context.DrawRectangle(fill, stroke, new Rect(handle.X - half, handle.Y - half, handleSize, handleSize));
        }
    }

    private void DrawPendingRedaction(DrawingContext context) {
        if (PendingRedactionArea is not Rect area || area.Width <= 0D || area.Height <= 0D) return;
        var fill = new SolidColorBrush(Color.FromArgb(58, 220, 38, 38));
        var stroke = new Pen(new SolidColorBrush(Color.FromArgb(235, 220, 38, 38)), 2D);
        context.DrawRectangle(fill, stroke, area);

        double corner = Math.Min(10D, Math.Min(area.Width, area.Height) / 3D);
        if (corner <= 1D) return;
        context.DrawLine(stroke, area.TopLeft, area.TopLeft + new Vector(corner, corner));
        context.DrawLine(stroke, area.BottomRight, area.BottomRight - new Vector(corner, corner));
    }

    private void DrawEditorPreview(DrawingContext context) {
        if (!_editing || _editorPath.Count == 0) return;
        var stroke = new Pen(new SolidColorBrush(Color.FromArgb(220, 220, 38, 38)), 1.5D);
        if (EditorTool == PdfEditorTool.Ink || EditorTool == PdfEditorTool.Line) {
            for (int index = 1; index < _editorPath.Count; index++) {
                context.DrawLine(stroke, _editorPath[index - 1], _editorPath[index]);
            }
            return;
        }
        Rect bounds = GetEditorBounds();
        var fill = new SolidColorBrush(Color.FromArgb(EditorTool == PdfEditorTool.Redact ? (byte)90 : (byte)28, 220, 38, 38));
        context.DrawRectangle(fill, stroke, bounds);
    }

    private void EmitEditorGesture() {
        if (Scene is null || _editorPath.Count == 0) return;
        Rect bounds = EnsureUsableBounds(GetEditorBounds(), EditorTool);
        EditorGestureCompleted?.Invoke(new PdfEditorGesture(
            Scene.PageNumber,
            bounds.Left,
            bounds.Top,
            bounds.Right,
            bounds.Bottom,
            _editorPath.Select(static point => new PdfEditorVisualPoint(point.X, point.Y)).ToArray()));
    }

    private Rect GetEditorBounds() {
        double left = _editorPath.Min(static point => point.X);
        double top = _editorPath.Min(static point => point.Y);
        double right = _editorPath.Max(static point => point.X);
        double bottom = _editorPath.Max(static point => point.Y);
        return new Rect(left, top, Math.Max(0D, right - left), Math.Max(0D, bottom - top));
    }

    private static Rect EnsureUsableBounds(Rect bounds, PdfEditorTool tool) {
        if (bounds.Width >= 4D && bounds.Height >= 4D) return bounds;
        (double Width, double Height) size = tool switch {
            PdfEditorTool.Note => (18D, 18D),
            PdfEditorTool.AddText => (160D, 30D),
            PdfEditorTool.Stamp => (144D, 48D),
            PdfEditorTool.Link => (140D, 24D),
            PdfEditorTool.AddImage => (160D, 100D),
            _ => (48D, 28D)
        };
        return new Rect(bounds.X, bounds.Y, size.Width, size.Height);
    }

    private static double Distance(Point left, Point right) {
        double x = right.X - left.X;
        double y = right.Y - left.Y;
        return Math.Sqrt((x * x) + (y * y));
    }

    private void ResetPointerState() {
        _selecting = false;
        _editing = false;
        _selectionStart = null;
        _selectionEnd = null;
        _hoverRegion = null;
        _editorPath.Clear();
    }

    private Point ToPagePoint(Point controlPoint) {
        PdfPageScene? scene = Scene;
        if (scene is null) return default;
        double scaleX = Bounds.Width / Math.Max(1D, scene.Drawing.Width);
        double scaleY = Bounds.Height / Math.Max(1D, scene.Drawing.Height);
        return new Point(controlPoint.X / Math.Max(scaleX, 0.000001D), controlPoint.Y / Math.Max(scaleY, 0.000001D));
    }
}
