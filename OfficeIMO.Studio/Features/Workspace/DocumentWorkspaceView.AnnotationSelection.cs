using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.VisualTree;
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
        var selections = PdfAnnotationMarqueeProjection.Project(reader, viewport, e, _marqueePages);
        if (selections is null) return;
        if (e.Completed) _document.OnPageAnnotationsSelected(new(selections, e.Additive, Toggle: false));
        e.Handled = true;
    }
}
