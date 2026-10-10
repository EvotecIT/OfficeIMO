using Avalonia;
using Avalonia.Interactivity;
using Avalonia.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

internal sealed class PdfAnnotationMarqueeEventArgs(PdfPageCanvas canvas, Point start, Point end,
    bool additive, bool completed, bool cancelled) : RoutedEventArgs(PdfPageCanvas.AnnotationMarqueeEvent) {
    internal PdfPageCanvas Canvas { get; } = canvas;
    internal Point Start { get; } = start;
    internal Point End { get; } = end;
    internal bool Additive { get; } = additive;
    internal bool Completed { get; } = completed;
    internal bool Cancelled { get; } = cancelled;
}

public sealed partial class PdfPageCanvas {
    internal static readonly RoutedEvent<PdfAnnotationMarqueeEventArgs> AnnotationMarqueeEvent =
        RoutedEvent.Register<PdfPageCanvas, PdfAnnotationMarqueeEventArgs>("AnnotationMarquee", RoutingStrategies.Bubble);

    private Rect? _annotationMarqueePreview;
    private IPointer? _annotationMarqueePointer;

    internal Rect? AnnotationMarqueePreview {
        get => _annotationMarqueePreview;
        set { _annotationMarqueePreview = value; InvalidateVisual(); }
    }

    private PdfAnnotationMarqueeEventArgs RaiseAnnotationMarquee(bool completed = false, bool cancelled = false) {
        var request = new PdfAnnotationMarqueeEventArgs(this, _selectionStart ?? default, _selectionEnd ?? default,
            _additiveAnnotationSelection, completed, cancelled);
        RaiseEvent(request);
        return request;
    }

    private void CancelAnnotationMarquee() {
        RaiseAnnotationMarquee(cancelled: true);
        _selecting = false;
        _selectionStart = null; _selectionEnd = null;
        var pointer = _annotationMarqueePointer;
        _annotationMarqueePointer = null;
        pointer?.Capture(null);
        InvalidateVisual();
    }

    internal IReadOnlyList<PdfEditorSelection> GetAnnotationsInControlRectangle(Rect rectangle) {
        if (Scene?.Interactions is not { } interactions || rectangle.Width <= 0 || rectangle.Height <= 0) return [];
        var pageRectangle = new Rect(ToPagePoint(rectangle.TopLeft), ToPagePoint(rectangle.BottomRight)).Normalize();
        return interactions.Regions.Where(region => region.Kind == PdfInteractionKind.Annotation && region.ObjectNumber.HasValue &&
            pageRectangle.Intersects(new Rect(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height)))
            .Select(region => CreateSelection(Scene.PageNumber, region)).ToArray();
    }

    internal Rect MapControlRectangleToPage(Rect rectangle) =>
        new Rect(ToPagePoint(rectangle.TopLeft), ToPagePoint(rectangle.BottomRight)).Normalize();
}
