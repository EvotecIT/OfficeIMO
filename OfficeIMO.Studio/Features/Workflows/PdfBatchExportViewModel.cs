using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Thin local-folder surface for the shared durable batch runner.</summary>
public sealed partial class PdfBatchExportViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickFolder;
    private readonly IOfficeWorkflowPublicationGuard? _guard;
    private readonly Func<bool> _otherWorkBusy;
    private readonly IStudioLocalizer _localizer;
    private readonly IOfficeWorkflowRunner _runner;
    private readonly StudioJobHistory? _jobs;
    private CancellationTokenSource? _cancellation;
    private bool _disposed;
    internal PdfBatchExportViewModel(Func<CancellationToken, Task<string?>> pickFolder, IOfficeWorkflowPublicationGuard? guard,
        Func<bool> otherWorkBusy, IStudioLocalizer localizer, IOfficeWorkflowRunner runner, StudioJobHistory? jobs,
        Func<Infrastructure.IStudioFileDialogs>? fileDialogs = null) {
        _pickFolder = pickFolder; _guard = guard; _otherWorkBusy = otherWorkBusy;
        _localizer = localizer;
        _runner = runner; _jobs = jobs;
        _fileDialogs = fileDialogs ?? (() => Infrastructure.NoStudioFileDialogs.Instance);
        Status = UnavailableReason ?? _localizer.GetOrDefault("Conversion.BatchExport.Ready", "Choose source and PDF output folders. Add optional checkpoints to resume mixed-format document export.");
    }
    [ObservableProperty] private string _inputDirectory = string.Empty;
    [ObservableProperty] private string _outputDirectory = string.Empty;
    [ObservableProperty] private string _checkpointDirectory = string.Empty;
    [ObservableProperty] private bool _allowLegacyImportLoss;
    [ObservableProperty] private bool _retryFailed;
    [ObservableProperty] private string _textEncoding = string.Empty;
    [ObservableProperty] private decimal _tabSize = 8;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _status = string.Empty;
    /// <summary>The batch runner walks and resumes whole folders by path, which sandboxed and mobile builds cannot reach.</summary>
    public string? UnavailableReason => Infrastructure.StudioDistributionPolicy.RequiresPickedLocations
        ? _localizer.GetOrDefault("Conversion.BatchExport.Unavailable", "Folder PDF export needs ongoing access to whole folders, which this edition of Studio does not have. Use Add files and Run pending to convert several documents.")
        : null;
    public bool IsAvailable => UnavailableReason is null;
    public bool CanEdit => IsAvailable && !_disposed && !IsBusy && !_otherWorkBusy();
    internal void RefreshHostState() {
        OnPropertyChanged(nameof(CanEdit)); ChooseFolderCommand.NotifyCanExecuteChanged(); RunCommand.NotifyCanExecuteChanged();
        RefreshSavedJobCommands();
    }
    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanEdit));
        ChooseFolderCommand.NotifyCanExecuteChanged(); RunCommand.NotifyCanExecuteChanged(); CancelCommand.NotifyCanExecuteChanged();
        RefreshSavedJobCommands();
    }
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task ChooseFolderAsync(string target, CancellationToken token) {
        string? folder = await _pickFolder(token).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(folder) || IsBusy) return;
        string? local = OfficeStorageIdentity.GetLocalPath(folder);
        if (local == null || !Path.IsPathFullyQualified(local)) { Status = _localizer.GetOrDefault("Conversion.BatchExport.LocalOnly", "Batch processing requires local filesystem folders."); return; }
        switch (target) { case "input": InputDirectory = local; break; case "output": OutputDirectory = local; break; case "state": CheckpointDirectory = local; break; }
    }
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task RunAsync() {
        OfficeConversionBatchRequest request;
        OfficeConversionJob definition;
        try {
            definition = CreateCurrentJob();
            request = definition.CreateRequest(_jobBaseDirectory);
        } catch (Exception error) when (error is IOException or ArgumentException or InvalidOperationException or NotSupportedException) {
            Status = error.Message;
            return;
        }
        using var cancellation = new CancellationTokenSource();
        _cancellation = cancellation;
        IsBusy = true;
        var progress = new BatchProgress();
        StudioJobRecord? job = null;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        timer.Tick += (_, _) => {
            Status = _localizer.FormatOrDefault("Conversion.BatchExport.Progress", "Completed {0}; failed {1}. Completed files survive cancellation.", progress.Completed, progress.Failed);
            job?.Report(new("folder-export", "execute", Status, 0, 0));
        };
        try {
            job = _jobs?.Start(_localizer.GetOrDefault("Conversion.BatchExport.Title", "Folder PDF export"),
                request.InputDirectory!, request.OutputDirectory, cancellation.Cancel, batch: true);
            Status = _localizer.GetOrDefault("Jobs.Queued", "Queued");
            using IDisposable? permit = _jobs is null ? null : await _jobs.EnterAsync(cancellation.Token).ConfigureAwait(true);
            Status = _localizer.GetOrDefault("Conversion.BatchExport.Running", "Exporting documents. Checkpointed jobs retain completed files individually.");
            job?.Report(new("folder-export", "execute", Status, 0, 0));
            timer.Start();
            OfficeConversionBatchResult result = await OfficeConversionJobRunner.RunAsync(definition, _jobBaseDirectory, _runner,
                progress: progress, cancellationToken: cancellation.Token, publicationGuard: _guard).ConfigureAwait(true);
            _hasRun = true;
            OnPropertyChanged(nameof(RunLabel));
            Status = _localizer.FormatOrDefault(result.FailureSummary is not null ? "Conversion.BatchExport.StoppedResult"
                    : result.Cancelled ? "Conversion.BatchExport.CancelledResult" : "Conversion.BatchExport.FinishedResult",
                result.FailureSummary is not null ? "Stopped: {0} completed, {1} reused, {2} failed, {3} skipped. Completed outputs remain available."
                    : result.Cancelled ? "Cancelled: {0} completed, {1} reused, {2} failed, {3} skipped." : "Finished: {0} completed, {1} reused, {2} failed, {3} skipped.",
                result.Completed, result.Reused, result.Failed, result.Skipped);
            if (result.FailureSummary is { } batchFailure) Status += " " + batchFailure;
            if (progress.FirstFailure is { } failure) Status += " " + failure;
            OfficeWorkflowStatus outcome = progress.HasUnconfirmed ? OfficeWorkflowStatus.Unconfirmed
                : result.FailureSummary is not null ? OfficeWorkflowStatus.Failed
                : result.Cancelled ? OfficeWorkflowStatus.Cancelled : result.Failed > 0 ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed;
            bool hasVerifiedOutput = result.Completed > 0 || result.Reused > 0;
            job?.CompleteBatch(outcome, hasVerifiedOutput ? request.OutputDirectory : null, Status, [], hasVerifiedOutput, isDirectoryOutput: true);
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            Status = _localizer.GetOrDefault("Conversion.BatchExport.Cancelled", "Batch cancelled. Completed files remain available.");
            job?.CompleteBatch(OfficeWorkflowStatus.Cancelled, progress.Completed > 0 ? request.OutputDirectory : null, Status, [], progress.Completed > 0, isDirectoryOutput: true);
        }
        catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            Status = error.Message;
            job?.CompleteBatch(OfficeWorkflowStatus.Failed, progress.Completed > 0 ? request.OutputDirectory : null, Status, [], progress.Completed > 0, isDirectoryOutput: true);
        }
        finally { timer.Stop(); _cancellation = null; IsBusy = false; }
    }
    [RelayCommand(CanExecute = nameof(IsBusy))]
    private void Cancel() => _cancellation?.Cancel();
    /// <inheritdoc />
    public void Dispose() { if (_disposed) return; _disposed = true; _cancellation?.Cancel(); RefreshHostState(); }
    private sealed class BatchProgress : IProgress<OfficeConversionBatchItemResult> {
        private long _completed, _failed;
        private string? _firstFailure;
        private int _unconfirmed;
        public string? FirstFailure => Volatile.Read(ref _firstFailure);
        public long Completed => Interlocked.Read(ref _completed);
        public long Failed => Interlocked.Read(ref _failed);
        public bool HasUnconfirmed => Volatile.Read(ref _unconfirmed) != 0;
        public void Report(OfficeConversionBatchItemResult item) {
            if (item.Status == OfficeWorkflowStatus.Completed) Interlocked.Increment(ref _completed);
            else if (item.Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed) {
                if (item.Status == OfficeWorkflowStatus.Unconfirmed) Interlocked.Exchange(ref _unconfirmed, 1);
                Interlocked.Increment(ref _failed);
                string message = Path.GetFileName(item.InputPath) + ": " + item.Summary;
                Interlocked.CompareExchange(ref _firstFailure, message.Length > 1024 ? message[..1024] : message, null);
            }
        }
    }
}
