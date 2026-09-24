using Avalonia.Controls;
using System.ComponentModel;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Transformation;
using Avalonia.Threading;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private static readonly TimeSpan ToastLifetime = TimeSpan.FromSeconds(6);
    private DispatcherTimer? _toastTimer;
    private bool _toastDismissed;
    private MainWindowViewModel? _toastDocument;

    private Features.Workflows.StudioJobHistory? _jobHistory;
    private Features.Workflows.StudioJobRecord? _jobToastRecord;
    private DispatcherTimer? _jobToastTimer;

    private void AttachJobToasts(MainWindowViewModel document) {
        if (ReferenceEquals(_jobHistory, document.Jobs.History)) return;
        if (_jobHistory is not null) {
            ((System.Collections.Specialized.INotifyCollectionChanged)_jobHistory.Entries).CollectionChanged -= OnJobEntriesChanged;
            foreach (var entry in _jobHistory.Entries) entry.PropertyChanged -= OnJobRecordChanged;
        }
        _jobHistory = document.Jobs.History;
        ((System.Collections.Specialized.INotifyCollectionChanged)_jobHistory.Entries).CollectionChanged += OnJobEntriesChanged;
        foreach (var entry in _jobHistory.Entries) entry.PropertyChanged += OnJobRecordChanged;
    }

    private void OnJobEntriesChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e) {
        foreach (var entry in e.OldItems?.OfType<Features.Workflows.StudioJobRecord>() ?? []) entry.PropertyChanged -= OnJobRecordChanged;
        foreach (var entry in e.NewItems?.OfType<Features.Workflows.StudioJobRecord>() ?? []) entry.PropertyChanged += OnJobRecordChanged;
    }

    private void OnJobRecordChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(Features.Workflows.StudioJobRecord.IsActive) ||
            sender is not Features.Workflows.StudioJobRecord { IsActive: false } record) return;
        Dispatcher.UIThread.Post(() => ShowJobToast(record));
    }

    private void ShowJobToast(Features.Workflows.StudioJobRecord record) {
        if (_windowClosed || ViewModel.IsJobsMode) return;
        _jobToastRecord = record;
        JobToastTitle.Text = record.Title;
        JobToastDetail.Text = string.IsNullOrWhiteSpace(record.Summary) ? record.Status : record.Summary;
        JobToastOpenButton.IsVisible = record.HasOutput;
        // The badge follows the outcome: a check only for verified success.
        (string icon, string tone) = record.IsSucceeded ? ("IconCheck", "Success")
            : record.IsFailed ? ("IconWarning", "Error") : ("IconInfo", "Accent");
        if (this.TryFindResource(icon, out object? geometry) && geometry is Avalonia.Media.Geometry data) JobToastIcon.Data = data;
        if (this.TryFindResource("Studio" + tone + "Brush", ActualThemeVariant, out object? foreground) && foreground is Avalonia.Media.IBrush brush) JobToastIcon.Foreground = brush;
        if (this.TryFindResource("Studio" + tone + "SoftBrush", ActualThemeVariant, out object? background) && background is Avalonia.Media.IBrush soft) JobToastBadge.Background = soft;
        JobToast.IsVisible = true;
        Dispatcher.UIThread.Post(() => {
            JobToast.Opacity = 1;
            JobToast.RenderTransform = TransformOperations.Parse("translateY(0px)");
        }, DispatcherPriority.Render);
        _jobToastTimer?.Stop();
        _jobToastTimer ??= new DispatcherTimer(TimeSpan.FromSeconds(10), DispatcherPriority.Background, (_, _) => HideJobToast());
        _jobToastTimer.Start();
    }

    private void HideJobToast() {
        _jobToastTimer?.Stop();
        if (!JobToast.IsVisible) return;
        JobToast.Opacity = 0;
        JobToast.RenderTransform = TransformOperations.Parse("translateY(12px)");
        DispatcherTimer.RunOnce(() => { if (JobToast.Opacity == 0) JobToast.IsVisible = false; }, TimeSpan.FromMilliseconds(220));
    }

    private void OnDismissJobToastClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => HideJobToast();

    private async void OnJobToastOpenClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        Features.Workflows.StudioJobRecord? record = _jobToastRecord;
        HideJobToast();
        if (record is not null && ViewModel.Jobs.OpenOutputCommand.CanExecute(record))
            await ViewModel.Jobs.OpenOutputCommand.ExecuteAsync(record);
    }

    private void OnJobToastViewClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        HideJobToast();
        ViewModel.Commands["Jobs"].Execute(null);
    }

    private void AttachOperationToast(MainWindowViewModel document) {
        AttachJobToasts(document);
        _toastDocument = document;
        document.PropertyChanged += OnToastDocumentChanged;
        OperationToast.PointerEntered -= OnToastPointerEntered;
        OperationToast.PointerExited -= OnToastPointerExited;
        OperationToast.PointerEntered += OnToastPointerEntered;
        OperationToast.PointerExited += OnToastPointerExited;
        _toastDismissed = false;
        UpdateOperationToast(restartTimer: true);
    }

    private void DetachOperationToast(MainWindowViewModel document) {
        document.PropertyChanged -= OnToastDocumentChanged;
        if (ReferenceEquals(_toastDocument, document)) _toastDocument = null;
    }

    private void OnToastDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        switch (e.PropertyName) {
            case nameof(MainWindowViewModel.OperationStatus):
                _toastDismissed = false;
                UpdateOperationToast(restartTimer: true);
                break;
            case nameof(MainWindowViewModel.IsWorkspaceBusy):
                UpdateOperationToast(restartTimer: true);
                break;
            case nameof(MainWindowViewModel.HasVisibleOperationStatus):
            case nameof(MainWindowViewModel.HasRecovery):
            case nameof(MainWindowViewModel.CanUndo):
            case nameof(MainWindowViewModel.WorkspaceMode):
                UpdateOperationToast(restartTimer: false);
                break;
        }
    }

    // Results fade in over the canvas instead of pushing the workspace down. Progress and recovery
    // choices stay until they resolve; plain results leave after a few seconds unless hovered.
    private void UpdateOperationToast(bool restartTimer) {
        MainWindowViewModel? document = _toastDocument;
        // Search progress already lives in the search pane; repeating it as a notice only covers the page.
        bool show = document is { HasVisibleOperationStatus: true, IsQuietOperationStatus: false } && !_toastDismissed;
        bool persistent = document is { IsWorkspaceBusy: true } or { HasRecovery: true };
        ToastUndoButton.IsVisible = document is { CanUndo: true, IsWorkspaceBusy: false, HasRecovery: false, IsPdfWorkspaceMode: true };
        if (show) ShowToast(); else HideToast();
        if (!restartTimer) return;
        _toastTimer?.Stop();
        if (show && !persistent) {
            _toastTimer ??= new DispatcherTimer(ToastLifetime, DispatcherPriority.Background, (_, _) => {
                _toastTimer?.Stop();
                if (_toastDocument is { IsWorkspaceBusy: false, HasRecovery: false }) {
                    _toastDismissed = true;
                    HideToast();
                }
            });
            _toastTimer.Interval = ToastLifetime;
            _toastTimer.Start();
        }
    }

    private void ShowToast() {
        if (!OperationToast.IsVisible) {
            OperationToast.IsVisible = true;
            Dispatcher.UIThread.Post(() => {
                if (!OperationToast.IsVisible) return;
                OperationToast.Opacity = 1;
                OperationToast.RenderTransform = TransformOperations.Parse("translateY(0px)");
            }, DispatcherPriority.Render);
        } else {
            OperationToast.Opacity = 1;
            OperationToast.RenderTransform = TransformOperations.Parse("translateY(0px)");
        }
    }

    private void HideToast() {
        if (!OperationToast.IsVisible) return;
        OperationToast.Opacity = 0;
        OperationToast.RenderTransform = TransformOperations.Parse("translateY(12px)");
        DispatcherTimer.RunOnce(() => {
            if (OperationToast.Opacity == 0) OperationToast.IsVisible = false;
        }, TimeSpan.FromMilliseconds(220));
    }

    private void OnToastPointerEntered(object? sender, PointerEventArgs e) => _toastTimer?.Stop();

    private void OnToastPointerExited(object? sender, PointerEventArgs e) {
        if (_toastDocument is { IsWorkspaceBusy: false, HasRecovery: false } && OperationToast.IsVisible) _toastTimer?.Start();
    }

    private void OnDismissToastClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        _toastDismissed = true;
        _toastTimer?.Stop();
        HideToast();
    }

    // Single-key tool shortcuts while annotating; letters match the tool names shown in tooltips.
    private bool TrySelectToolShortcut(Key key) {
        if (!ViewModel.IsPdfWorkspaceMode || !ViewModel.HasDocument || ViewModel.IsFocusReading) return false;
        string? tool = ViewModel.IsAnnotateDocumentMode ? key switch {
            Key.V => "Select",
            Key.N => "Note",
            Key.T => "FreeText",
            Key.H => "Highlight",
            Key.U => "Underline",
            Key.S => "StrikeOut",
            Key.R => "Rectangle",
            Key.E => "Ellipse",
            Key.L => "Line",
            Key.I => "Ink",
            Key.Q => "Squiggly",
            Key.P => "Polygon",
            _ => null
        } : ViewModel.IsEditDocumentMode ? key switch {
            Key.V => "Select",
            Key.T => "AddText",
            Key.K => "Link",
            _ => null
        } : null;
        if (tool is null || !ViewModel.SelectEditorToolCommand.CanExecute(tool)) return false;
        ViewModel.SelectEditorToolCommand.Execute(tool);
        return true;
    }
}
