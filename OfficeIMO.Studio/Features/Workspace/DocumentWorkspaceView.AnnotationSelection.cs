using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView {
    private readonly HashSet<PdfPageCanvas> _marqueePages = [];

    private void InitializeAnnotationSelectionInput() {
        AddHandler(PdfPageCanvas.AnnotationMarqueeEvent, OnAnnotationMarquee);
        DetachedFromVisualTree += (_, _) => ClearAnnotationMarquee();
        DataContextChanged += (_, _) => ClearAnnotationMarquee();
    }

    private void ClearAnnotationMarquee() {
        foreach (var canvas in _marqueePages) canvas.AnnotationMarqueePreview = null;
        _marqueePages.Clear();
    }

    // Page canvases keep pointer capture. The host only projects that rectangle onto
    // realized, visible pages in the same viewport, including continuous and grid layouts.
    private void OnAnnotationMarquee(object? sender, PdfAnnotationMarqueeEventArgs e) {
        ClearAnnotationMarquee();
        if (e.Cancelled) return;
        var origin = e.Canvas;
        var reader = origin.FindAncestorOfType<ListBox>();
        var viewport = origin.FindAncestorOfType<ScrollContentPresenter>();
        if (_document is null || reader is null || viewport is null ||
            (reader != PagesList && reader != GridPagesList) || !origin.IsEffectivelyVisible) return;
        Point? start = origin.TranslatePoint(e.Start, viewport), end = origin.TranslatePoint(e.End, viewport);
        if (start is null || end is null) return;
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
            if (e.Completed) selections.AddRange(canvas.GetAnnotationsInControlRectangle(local));
            else {
                canvas.AnnotationMarqueePreview = canvas.MapControlRectangleToPage(local);
                _marqueePages.Add(canvas);
            }
        }
        if (e.Completed) _document.OnPageAnnotationsSelected(new(selections, e.Additive, Toggle: false));
        e.Handled = true;
    }
}
