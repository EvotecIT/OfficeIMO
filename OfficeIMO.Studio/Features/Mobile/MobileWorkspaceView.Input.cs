using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Platform;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private IInputPane? _inputPane;
    private double _keyboardOverlap;

    private void InitializeTouchInput() {
        _pinchRecognizer.Started += () => _suppressPinchContext = true;
        PageScroll.GestureRecognizers.Add(_pinchRecognizer);
        PageScroll.DetachedFromVisualTree += (_, _) => CancelPinch();
        PageScroll.AddHandler(PointerPressedEvent, (_, _) => {
            if (_pinchStartZoom is null && !_ignorePinch) _suppressPinchContext = false;
        }, RoutingStrategies.Tunnel);
        PageScroll.AddHandler(ContextRequestedEvent, (sender, e) => {
            // A holding gesture can reach the page after its touches have been taken by the pinch recognizer.
            if (_suppressPinchContext && e.TryGetPosition(PageScroll, out _)) e.Handled = true;
        }, RoutingStrategies.Tunnel);
        PageScroll.AddHandler(InputElement.PinchEvent, OnPinch);
        PageScroll.AddHandler(InputElement.PinchEndedEvent, (_, _) => EndPinch());
        AttachedToVisualTree += (_, _) => {
            _inputPane = TopLevel.GetTopLevel(this)?.InputPane;
            if (_inputPane is not null) _inputPane.StateChanged += OnInputPaneChanged;
        };
        DetachedFromVisualTree += (_, _) => {
            if (_inputPane is not null) _inputPane.StateChanged -= OnInputPaneChanged;
            _inputPane = null;
            CancelPinch();
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
        _keyboardOverlap = overlap;
        ApplyKeyboardAvoidance();
        UpdateDialogLayout();
    }
    private void ApplyKeyboardAvoidance() {
        SheetScrim.Margin = new Thickness(0, 0, 0, _keyboardOverlap);
        DialogScrim.Margin = new Thickness(0, 0, 0, _keyboardOverlap);
        CommandPalette.Margin = new Thickness(0, 0, 0, _keyboardOverlap);
        AssistantHost.Margin = new Thickness(0, 0, 0, SheetScrim.IsVisible || DialogScrim.IsVisible ? 0 : _keyboardOverlap);
    }
}
