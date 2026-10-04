using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private Action<object?>? _dismissDialog;
    private double _dialogWidth;
    private double _dialogHeight;
    private double _dialogMinimumContentHeight;
    private bool _initializing;

    internal void SetInitializing(bool value) {
        _initializing = value;
        ApplicationNavigation.IsEnabled = !value && !DialogScrim.IsVisible;
    }

    /// <summary>Presents the shared review inside the application and cancels when its host detaches.</summary>
    internal async Task<T?> ShowDialogAsync<T>(StudioDialogContent content) {
        if (_dismissDialog is not null || SheetScrim.IsVisible || !this.IsAttachedToVisualTree()) {
            content.Dismissed();
            return default;
        }
        var completion = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var previousFocus = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() as Control;
        _dialogWidth = double.IsFinite(content.Width) ? content.Width : 560;
        double preferredHeight = double.IsFinite(content.Height) ? content.Height : 360;
        _dialogMinimumContentHeight = Math.Min(content.MinHeight, 320);
        DialogHeader.IsVisible = !content.HasContentHeading;
        _dialogHeight = preferredHeight + (DialogHeader.IsVisible ? 64 : 0);
        content.ClearValue(WidthProperty);
        content.ClearValue(HeightProperty);
        content.ClearValue(MinWidthProperty);
        content.ClearValue(MinHeightProperty);
        _dismissDialog = result => completion.TrySetResult(result);
        void Detached(object? sender, VisualTreeAttachmentEventArgs e) => completion.TrySetResult(null);
        content.CloseRequested += _dismissDialog;
        DetachedFromVisualTree += Detached;
        ApplicationNavigation.IsEnabled = false;
        DialogTitle.Text = content.Title;
        DialogContent.Content = content;
        if (_dialogMinimumContentHeight == 0) {
            content.Measure(new Size(Math.Max(1, Math.Min(_dialogWidth, Bounds.Width - 24)), double.PositiveInfinity));
            _dialogMinimumContentHeight = content.DesiredSize.Height;
            _dialogHeight = Math.Max(_dialogHeight, _dialogMinimumContentHeight + (DialogHeader.IsVisible ? 64 : 0));
        }
        DialogScrim.IsVisible = true;
        ApplyKeyboardAvoidance();
        UpdateDialogLayout();
        Dispatcher.UIThread.Post(() => {
            if (!completion.Task.IsCompleted) content.Presented(this);
        }, DispatcherPriority.Loaded);
        try {
            object? result = await completion.Task;
            return result is T value ? value : default;
        } finally {
            content.CloseRequested -= _dismissDialog;
            DetachedFromVisualTree -= Detached;
            _dismissDialog = null;
            DialogScrim.IsVisible = false;
            ApplyKeyboardAvoidance();
            DialogContent.Content = null;
            ApplicationNavigation.IsEnabled = !_initializing;
            content.Dismissed();
            previousFocus?.Focus();
        }
    }

    private void OnDismissDialogClick(object? sender, RoutedEventArgs e) => _dismissDialog?.Invoke(null);

    private void UpdateDialogLayout() {
        if (!DialogScrim.IsVisible) return;
        DialogCard.Width = Math.Max(1, Math.Min(_dialogWidth, Bounds.Width - 24));
        DialogCard.Height = Math.Max(1, Math.Min(_dialogHeight, Bounds.Height - DialogScrim.Margin.Bottom - 24));
        DialogContent.Height = Math.Max(_dialogMinimumContentHeight, DialogCard.Height - (DialogHeader.IsVisible ? 64 : 0));
    }
}
