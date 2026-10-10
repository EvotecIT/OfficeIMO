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
            Point? position = canvas.TranslatePoint(default, viewport);
            if (position is null) continue;
            Rect visible = new Rect(position.Value, canvas.Bounds.Size).Intersect(new Rect(viewport.Bounds.Size));
            Rect intersection = rectangle.Intersect(visible);
            if (intersection.Width <= 0 || intersection.Height <= 0) continue;
            var local = intersection.Translate(new Vector(-position.Value.X, -position.Value.Y));
            if (request.Completed) selections.AddRange(canvas.GetAnnotationsInControlRectangle(local));
            else {
                canvas.AnnotationMarqueePreview = canvas.MapControlRectangleToPage(local);
                previews.Add(canvas);
            }
        }
        return selections;
    }
}
