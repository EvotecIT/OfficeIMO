using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

internal sealed record PdfObjectTransformGesture(PdfEditorSelection Selection, PdfEditorVisualBounds Target);

public sealed partial class PdfPageCanvas {
    private PdfEditorSelection? _transformSelection;
    private Point _transformStart;
    private Rect _transformTarget;
    private int _transformHandle;
    private double? _alignmentX;
    private double? _alignmentY;
    internal event Action<PdfObjectTransformGesture>? ObjectTransformCompleted;

    private bool BeginObjectTransform(PointerPressedEventArgs e) {
        if (SelectedObject is not { Kind: PdfEditorSelectionKind.Image or PdfEditorSelectionKind.Annotation } selected ||
            SelectionMode is not (PdfEditorSelectionMode.PageContent or PdfEditorSelectionMode.Annotations)) return false;
        Point point = ToPagePoint(e.GetPosition(this));
        var bounds = new Rect(selected.Bounds.Left, selected.Bounds.Top, selected.Bounds.Width, selected.Bounds.Height);
        Point[] handles = HandlePoints(bounds);
        double tolerance = 7 * Math.Max(1, (Scene?.Drawing.Width ?? Bounds.Width) / Math.Max(1, Bounds.Width));
        int handle = Array.FindIndex(handles, candidate => Distance(point, candidate) <= tolerance);
        if (handle < 0 && !bounds.Contains(point)) return false;
        _transformSelection = selected;
        _transformStart = point;
        _transformTarget = bounds;
        _transformHandle = handle;
        e.Pointer.Capture(this);
        e.Handled = true;
        return true;
    }

    private bool UpdateObjectTransform(PointerEventArgs e) {
        if (_transformSelection is not { } selection || Scene is not { } scene) return false;
        Point point = ToPagePoint(e.GetPosition(this));
        var original = new Rect(selection.Bounds.Left, selection.Bounds.Top, selection.Bounds.Width, selection.Bounds.Height);
        var delta = point - _transformStart;
        Rect target = _transformHandle < 0 ? original.Translate(delta) : ResizeObjectBounds(original, delta,
            _transformHandle, selection.Kind == PdfEditorSelectionKind.Image);
        _alignmentX = null; _alignmentY = null;
        if (_transformHandle < 0) {
            // Align against the page and nearby object edges, in rendered page coordinates.
            double[] xGuides = new[] { 0D, scene.Drawing.Width / 2D, scene.Drawing.Width }
                .Concat(scene.Interactions?.Regions.Where(region => Math.Abs(region.Quad.Left - original.Left) > 0.01 || Math.Abs(region.Quad.Top - original.Top) > 0.01)
                    .SelectMany(region => new[] { region.Quad.Left, region.Quad.Right }) ?? []).Take(1000).ToArray();
            double[] yGuides = new[] { 0D, scene.Drawing.Height / 2D, scene.Drawing.Height }
                .Concat(scene.Interactions?.Regions.Where(region => Math.Abs(region.Quad.Left - original.Left) > 0.01 || Math.Abs(region.Quad.Top - original.Top) > 0.01)
                    .SelectMany(region => new[] { region.Quad.Top, region.Quad.Bottom }) ?? []).Take(1000).ToArray();
            (double xOffset, _alignmentX) = SnapOffset([target.Left, target.Center.X, target.Right], xGuides);
            (double yOffset, _alignmentY) = SnapOffset([target.Top, target.Center.Y, target.Bottom], yGuides);
            target = target.Translate(new Vector(xOffset, yOffset));
        }
        _transformTarget = target;
        e.Handled = true;
        InvalidateVisual();
        return true;
    }

    private bool CompleteObjectTransform(PointerReleasedEventArgs e) {
        if (_transformSelection is not { } selection) return false;
        UpdateObjectTransform(e);
        Rect target = _transformTarget;
        ResetObjectTransform();
        e.Pointer.Capture(null);
        e.Handled = true;
        if (Math.Abs(target.Left - selection.Bounds.Left) + Math.Abs(target.Top - selection.Bounds.Top) +
            Math.Abs(target.Width - selection.Bounds.Width) + Math.Abs(target.Height - selection.Bounds.Height) > 0.1) {
            ObjectTransformCompleted?.Invoke(new(selection, new(target.Left, target.Top, target.Right, target.Bottom)));
        }
        InvalidateVisual();
        return true;
    }

    private void ResetObjectTransform() { _transformSelection = null; _alignmentX = null; _alignmentY = null; }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e) {
        base.OnPointerCaptureLost(e);
        ResetObjectTransform();
        InvalidateVisual();
    }

    private void DrawObjectTransform(DrawingContext context) {
        if (_transformSelection is null || Scene is not { } scene) return;
        var pen = new Pen(Brushes.DodgerBlue, 1.5);
        context.DrawRectangle(new SolidColorBrush(Color.FromArgb(35, 30, 144, 255)), pen, _transformTarget);
        DrawSelectionHandles(context, _transformTarget, Colors.DodgerBlue);
        if (_alignmentX is double x) context.DrawLine(pen, new Point(x, 0), new Point(x, scene.Drawing.Height));
        if (_alignmentY is double y) context.DrawLine(pen, new Point(0, y), new Point(scene.Drawing.Width, y));
    }

    internal static Rect ResizeObjectBounds(Rect original, Vector delta, int handle, bool proportional) {
        bool left = handle is 0 or 6 or 7, right = handle is 2 or 3 or 4;
        bool top = handle is 0 or 1 or 2, bottom = handle is 4 or 5 or 6;
        double width = Math.Max(4, original.Width + (left ? -delta.X : right ? delta.X : 0));
        double height = Math.Max(4, original.Height + (top ? -delta.Y : bottom ? delta.Y : 0));
        if (proportional) {
            double scale = left || right ? width / original.Width : height / original.Height;
            if ((left || right) && (top || bottom) && Math.Abs(delta.Y) > Math.Abs(delta.X)) scale = height / original.Height;
            width = original.Width * scale; height = original.Height * scale;
        }
        double x = left ? original.Right - width : right ? original.Left : original.Center.X - width / 2;
        double y = top ? original.Bottom - height : bottom ? original.Top : original.Center.Y - height / 2;
        return new Rect(x, y, width, height);
    }

    private static (double Offset, double? Guide) SnapOffset(double[] anchors, double[] guides) {
        double offset = 4D; double? guide = null;
        foreach (double candidate in guides) foreach (double anchor in anchors) {
            if (Math.Abs(candidate - anchor) < Math.Abs(offset)) { offset = candidate - anchor; guide = candidate; }
        }
        return guide.HasValue ? (offset, guide) : (0D, null);
    }

    private static Point[] HandlePoints(Rect area) => [area.TopLeft, new(area.Center.X, area.Top), area.TopRight,
        new(area.Right, area.Center.Y), area.BottomRight, new(area.Center.X, area.Bottom), area.BottomLeft, new(area.Left, area.Center.Y)];
}
