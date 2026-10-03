using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Platform;
using Avalonia.Input;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private double? _pinchStartZoom;
    private IInputPane? _inputPane;

    private void InitializeTouchInput() {
        PageScroll.GestureRecognizers.Add(new PinchGestureRecognizer());
        PageScroll.AddHandler(InputElement.PinchEvent, (_, e) => {
            if (Document is not { } document) return;
            _pinchStartZoom ??= document.Zoom;
            document.SetTouchZoom(_pinchStartZoom.Value * e.Scale);
            e.Handled = true;
        });
        PageScroll.AddHandler(InputElement.PinchEndedEvent, (_, _) => _pinchStartZoom = null);
        AttachedToVisualTree += (_, _) => {
            _inputPane = TopLevel.GetTopLevel(this)?.InputPane;
            if (_inputPane is not null) _inputPane.StateChanged += OnInputPaneChanged;
        };
        DetachedFromVisualTree += (_, _) => {
            if (_inputPane is not null) _inputPane.StateChanged -= OnInputPaneChanged;
            _inputPane = null;
            _pinchStartZoom = null;
        };
    }

    private void OnInputPaneChanged(object? sender, InputPaneStateEventArgs e) {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel is null) return;
        var origin = this.TranslatePoint(default, topLevel);
        double overlap = e.NewState == InputPaneState.Open && origin is { } point &&
                         e.EndRect.Width >= topLevel.Bounds.Width * 0.8
            ? Math.Clamp(point.Y + Bounds.Height - e.EndRect.Top, 0, Bounds.Height)
            : 0;
        SheetScrim.Margin = new Thickness(0, 0, 0, overlap);
    }
}
