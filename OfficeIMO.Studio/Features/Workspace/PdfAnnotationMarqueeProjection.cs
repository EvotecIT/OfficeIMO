using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>Projects a captured selection rectangle across realized pages of one viewport.</summary>
internal static class PdfAnnotationMarqueeProjection {
    internal static IReadOnlyList<PdfEditorSelection>? Project(ListBox reader, ScrollContentPresenter viewport,
        PdfAnnotationMarqueeEventArgs request, ISet<PdfPageCanvas> previews) {
        Point? start = request.Canvas.TranslatePoint(request.Start, viewport), end = request.Canvas.TranslatePoint(request.End, viewport);
        if (start is null || end is null) return null;
        var rectangle = new Rect(start.Value, end.Value).Normalize().Intersect(new Rect(viewport.Bounds.Size));
        var selections = new List<PdfEditorSelection>();
        foreach (var canvas in reader.GetVisualDescendants().OfType<PdfPageCanvas>()) {
            if (!canvas.IsEffectivelyVisible || canvas.Scene is null || canvas.SelectionMode != PdfEditorSelectionMode.Annotations) continue;
            Rect? projected = ProjectRectangle(new Rect(canvas.Bounds.Size), canvas, viewport);
            if (projected is null) continue;
            Rect visible = projected.Value.Intersect(new Rect(viewport.Bounds.Size));
            Rect intersection = rectangle.Intersect(visible);
            if (intersection.Width <= 0 || intersection.Height <= 0) continue;
            Rect? local = ProjectRectangle(intersection, viewport, canvas);
            if (local is null) continue;
            if (request.Completed) selections.AddRange(canvas.GetAnnotationsInControlRectangle(local.Value));
            else {
                canvas.AnnotationMarqueePreview = canvas.MapControlRectangleToPage(local.Value);
                previews.Add(canvas);
            }
        }
        return selections;
    }

    private static Rect? ProjectRectangle(Rect rectangle, Visual source, Visual target) {
        Point? a = source.TranslatePoint(rectangle.TopLeft, target), b = source.TranslatePoint(rectangle.TopRight, target);
        Point? c = source.TranslatePoint(rectangle.BottomLeft, target), d = source.TranslatePoint(rectangle.BottomRight, target);
        if (a is null || b is null || c is null || d is null) return null;
        double left = Math.Min(Math.Min(a.Value.X, b.Value.X), Math.Min(c.Value.X, d.Value.X));
        double top = Math.Min(Math.Min(a.Value.Y, b.Value.Y), Math.Min(c.Value.Y, d.Value.Y));
        double right = Math.Max(Math.Max(a.Value.X, b.Value.X), Math.Max(c.Value.X, d.Value.X));
        double bottom = Math.Max(Math.Max(a.Value.Y, b.Value.Y), Math.Max(c.Value.Y, d.Value.Y));
        return new Rect(left, top, right - left, bottom - top);
    }
}
