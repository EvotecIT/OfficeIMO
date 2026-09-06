using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Navigation and non-destructive history actions for the shared Jobs surface.</summary>
public sealed partial class StudioJobsViewModel : ObservableObject, IDisposable {
    private readonly Func<string, CancellationToken, Task> _openOutput;
    private readonly Func<string, bool> _usesProvider;
    internal StudioJobsViewModel(StudioJobHistory history, Func<string, CancellationToken, Task> openOutput, Func<string, bool>? usesProvider = null) {
        History = history;
        _openOutput = openOutput;
        _usesProvider = usesProvider ?? (_ => false);
        History.PropertyChanged += OnHistoryChanged;
    }
    public StudioJobHistory History { get; }
    [ObservableProperty]
    private StudioJobRecord? _selectedJob;
    [ObservableProperty]
    private string? _actionError;
    private bool CanOpen(StudioJobRecord? job) => job?.HasOutput == true;
    public bool CanClear => History.CanClear;

    partial void OnSelectedJobChanged(StudioJobRecord? oldValue, StudioJobRecord? newValue) {
        if (oldValue is not null) oldValue.PropertyChanged -= OnSelectedChanged;
        if (newValue is not null) newValue.PropertyChanged += OnSelectedChanged;
        ActionError = null;
        OpenOutputCommand.NotifyCanExecuteChanged();
    }
    [RelayCommand(CanExecute = nameof(CanOpen))]
    private async Task OpenOutputAsync(StudioJobRecord? job, CancellationToken cancellationToken) {
        string? path = job?.OutputPath;
        if (!CanOpen(job) || path is null) return;
        try {
            ActionError = null;
            string? localPath = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(path);
            if (localPath is not null && !_usesProvider(path) && !File.Exists(localPath) && !Directory.Exists(localPath))
                throw new FileNotFoundException("The output is no longer available.", localPath);
            await _openOutput(path, cancellationToken).ConfigureAwait(true);
        } catch (Exception exception) {
            ActionError = exception.Message;
        }
    }
    [RelayCommand]
    private async Task OpenRecoveryAsync(StudioJobRecord? job, CancellationToken cancellationToken) {
        if (job?.Recovery is not { } recovery || History.RecoveryStore is not { } store) return;
        try {
            ActionError = null;
            await store.VerifyAsync(recovery, cancellationToken).ConfigureAwait(true);
            await _openOutput(recovery.FilePath, cancellationToken).ConfigureAwait(true);
        } catch (Exception exception) { ActionError = exception.Message; }
    }
    [RelayCommand]
    private void RequestDiscardRecovery(StudioJobRecord? job) {
        if (job?.HasRecovery == true) job.RecoveryDiscardRequested = true;
    }
    [RelayCommand]
    private void CancelDiscardRecovery(StudioJobRecord? job) {
        if (job is not null) job.RecoveryDiscardRequested = false;
    }
    [RelayCommand]
    private void ConfirmDiscardRecovery(StudioJobRecord? job) {
        if (job?.RecoveryDiscardRequested != true || job.Recovery is not { } recovery || History.RecoveryStore is not { } store) return;
        try {
            ActionError = null;
            store.Discard(recovery);
            job.RemoveRecovery(recovery);
            job.RecoveryDiscardRequested = false;
        } catch (Exception exception) { ActionError = exception.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanClear))]
    private void ClearFinished() {
        History.ClearFinished();
        if (SelectedJob is not null && !History.Entries.Contains(SelectedJob)) SelectedJob = null;
    }
    private void OnHistoryChanged(object? sender, PropertyChangedEventArgs args) {
        ClearFinishedCommand.NotifyCanExecuteChanged();
        OpenOutputCommand.NotifyCanExecuteChanged();
        if (SelectedJob is not null && !History.Entries.Contains(SelectedJob)) SelectedJob = null;
    }
    private void OnSelectedChanged(object? sender, PropertyChangedEventArgs args) => OpenOutputCommand.NotifyCanExecuteChanged();
    public void Dispose() {
        History.PropertyChanged -= OnHistoryChanged;
        if (SelectedJob is not null) SelectedJob.PropertyChanged -= OnSelectedChanged;
    }
}
