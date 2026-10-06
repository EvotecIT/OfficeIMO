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
        Func<bool> otherWorkBusy, IStudioLocalizer localizer, IOfficeWorkflowRunner runner, StudioJobHistory? jobs) {
        _pickFolder = pickFolder; _guard = guard; _otherWorkBusy = otherWorkBusy;
        _localizer = localizer;
        _runner = runner; _jobs = jobs;
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
    public string? UnavailableReason => OperatingSystem.IsIOS()
        ? "Folder PDF export needs persistent filesystem access and is not available on iPad or iPhone. Use Add files and Run pending to convert multiple documents from Files."
        : null;
    public bool IsAvailable => UnavailableReason is null;
    public bool CanEdit => IsAvailable && !_disposed && !IsBusy && !_otherWorkBusy();
    internal void RefreshHostState() {
        OnPropertyChanged(nameof(CanEdit)); ChooseFolderCommand.NotifyCanExecuteChanged(); RunCommand.NotifyCanExecuteChanged();
    }
    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanEdit));
        ChooseFolderCommand.NotifyCanExecuteChanged(); RunCommand.NotifyCanExecuteChanged(); CancelCommand.NotifyCanExecuteChanged();
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
        var request = new OfficeConversionBatchRequest { InputDirectory = InputDirectory, OutputDirectory = OutputDirectory,
            CheckpointDirectory = string.IsNullOrWhiteSpace(CheckpointDirectory) ? null : CheckpointDirectory, RetryFailed = RetryFailed,
            ConversionOptions = new OfficeWorkflowConversionOptions {
                LegacyDocLossPolicy = AllowLegacyImportLoss ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block,
                PlainText = new OfficeIMO.Pdf.PdfPlainTextOptions {
                    EncodingName = string.IsNullOrWhiteSpace(TextEncoding) ? null : TextEncoding, TabSize = (int)TabSize
                }
            } };
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
            OfficeConversionBatchResult result = await _runner.RunBatchAsync(request, progress, cancellation.Token, _guard).ConfigureAwait(true);
            Status = _localizer.FormatOrDefault(result.Cancelled ? "Conversion.BatchExport.CancelledResult" : "Conversion.BatchExport.FinishedResult",
                result.Cancelled ? "Cancelled: {0} completed, {1} reused, {2} failed, {3} skipped." : "Finished: {0} completed, {1} reused, {2} failed, {3} skipped.",
                result.Completed, result.Reused, result.Failed, result.Skipped);
            if (progress.FirstFailure is { } failure) Status += " " + failure;
            OfficeWorkflowStatus outcome = progress.HasUnconfirmed ? OfficeWorkflowStatus.Unconfirmed
                : result.Cancelled ? OfficeWorkflowStatus.Cancelled : result.Failed > 0 ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed;
            job?.CompleteBatch(outcome, result.Completed > 0 ? request.OutputDirectory : null, Status, [], result.Completed > 0, isDirectoryOutput: true);
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
