using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private readonly ReaderPinchGestureRecognizer _pinchRecognizer = new();
    private double? _pinchStartZoom;
    private PinchAnchor? _pinchAnchor;
    private (PinchAnchor Anchor, double Zoom)? _pendingPinchAdjustment;
    private bool _pinchAdjustmentQueued;
    private bool _ignorePinch;
    private bool _suppressPinchContext;

    private sealed record PinchAnchor(MainWindowViewModel Document, PdfPageViewModel Page,
        Border Paper, Point PageFraction, Point ViewportPoint);

    private sealed class ReaderPinchGestureRecognizer : PinchGestureRecognizer {
        private readonly HashSet<IPointer> _contacts = [];
        public bool HasContacts => _contacts.Count > 0;
        public event Action? Started;

        protected override void PointerPressed(PointerPressedEventArgs e) {
            if (e.Pointer.Type is PointerType.Touch or PointerType.Pen && _contacts.Count < 2 && _contacts.Add(e.Pointer) && _contacts.Count == 2)
                Started?.Invoke();
            base.PointerPressed(e);
        }

        protected override void PointerReleased(PointerReleasedEventArgs e) {
            _contacts.Remove(e.Pointer);
            base.PointerReleased(e);
        }

        protected override void PointerCaptureLost(IPointer pointer) {
            _contacts.Remove(pointer);
            base.PointerCaptureLost(pointer);
        }
    }

    private void OnPinch(object? sender, PinchEventArgs e) {
        if (_ignorePinch || TopLevel.GetTopLevel(PageScroll) is null || !PageScroll.IsEffectivelyVisible ||
            !double.IsFinite(e.Scale) || e.Scale <= 0 ||
            Document is not { SelectedPage: { } page } document) return;
        _suppressPinchContext = true;
        if (_pinchStartZoom is null) {
            _pinchStartZoom = document.Zoom;
            var view = PageScroll.GetVisualDescendants().OfType<PdfPageView>()
                .FirstOrDefault(view => ReferenceEquals(view.DataContext, page));
            if (view?.FindControl<Border>("PagePaper") is { Bounds.Width: > 0, Bounds.Height: > 0 } paper &&
                PageScroll.TranslatePoint(e.ScaleOrigin, paper) is { } point) {
                _pinchAnchor = new PinchAnchor(document, page, paper,
                    new Point(point.X / paper.Bounds.Width, point.Y / paper.Bounds.Height), e.ScaleOrigin);
            }
        }
        document.SetTouchZoom(_pinchStartZoom.Value * e.Scale);
        if (_pinchAnchor is { } anchor) {
            _pendingPinchAdjustment = (anchor, document.Zoom);
            if (!_pinchAdjustmentQueued) {
                _pinchAdjustmentQueued = true;
                // Several touch updates may arrive before layout. Adjust once using the latest page size.
                Dispatcher.UIThread.Post(ApplyPinchAnchor, DispatcherPriority.Loaded);
            }
        }
        e.Handled = true;
    }

    private void ApplyPinchAnchor() {
        _pinchAdjustmentQueued = false;
        var pending = _pendingPinchAdjustment;
        _pendingPinchAdjustment = null;
        if (pending is not { } adjustment) return;
        PinchAnchor anchor = adjustment.Anchor;
        if (!ReferenceEquals(Document, anchor.Document) || !ReferenceEquals(Document.SelectedPage, anchor.Page) ||
            !ReferenceEquals(anchor.Paper.DataContext, anchor.Page) ||
            !anchor.Paper.GetVisualAncestors().Contains(PageScroll) ||
            Math.Abs(anchor.Document.Zoom - adjustment.Zoom) > 0.001) return;
        var pagePoint = new Point(anchor.PageFraction.X * anchor.Paper.Bounds.Width,
            anchor.PageFraction.Y * anchor.Paper.Bounds.Height);
        if (anchor.Paper.TranslatePoint(pagePoint, PageScroll) is not { } current) return;
        // ScrollViewer clamps the offset at page edges and recenters pages smaller than the viewport.
        PageScroll.Offset += new Vector(current.X - anchor.ViewportPoint.X, current.Y - anchor.ViewportPoint.Y);
    }

    private void EndPinch() {
        _pinchStartZoom = null;
        _pinchAnchor = null;
        _ignorePinch = _ignorePinch && _pinchRecognizer.HasContacts;
        // Keep the final queued adjustment even when the fingers leave before the next layout.
    }

    private void CancelPinch() {
        // The recognizer captures contacts before the first scale update. Ignore all old
        // contacts through release, including a replacement finger while one remains down.
        _ignorePinch = _ignorePinch || _pinchStartZoom is not null || _pinchRecognizer.HasContacts;
        _pinchStartZoom = null;
        _pinchAnchor = null;
        _pendingPinchAdjustment = null;
    }
}
