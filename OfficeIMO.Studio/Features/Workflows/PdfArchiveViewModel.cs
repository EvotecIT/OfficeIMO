using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Thin local-folder surface for the shared durable archive runner.</summary>
public sealed partial class PdfArchiveViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickFolder;
    private readonly IOfficeWorkflowPublicationGuard? _guard;
    private readonly Func<bool> _otherWorkBusy;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _cancellation;
    internal PdfArchiveViewModel(Func<CancellationToken, Task<string?>> pickFolder, IOfficeWorkflowPublicationGuard? guard, Func<bool> otherWorkBusy, IStudioLocalizer localizer) {
        _pickFolder = pickFolder; _guard = guard; _otherWorkBusy = otherWorkBusy;
        _localizer = localizer;
        Status = _localizer.GetOrDefault("Conversion.Archive.Ready", "Choose separate source, PDF output and checkpoint folders. DOC, DOCX and TXT files are discovered recursively.");
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
    public bool CanEdit => !IsBusy && !_otherWorkBusy();
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
        if (local == null || !Path.IsPathFullyQualified(local)) { Status = _localizer.GetOrDefault("Conversion.Archive.LocalOnly", "Archive processing requires local filesystem folders."); return; }
        switch (target) { case "input": InputDirectory = local; break; case "output": OutputDirectory = local; break; case "state": CheckpointDirectory = local; break; }
    }
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task RunAsync() {
        var request = new OfficePdfArchiveRequest { InputDirectory = InputDirectory, OutputDirectory = OutputDirectory,
            CheckpointDirectory = CheckpointDirectory, AllowLegacyImportLoss = AllowLegacyImportLoss, RetryFailed = RetryFailed,
            TextEncoding = string.IsNullOrWhiteSpace(TextEncoding) ? null : TextEncoding, TabSize = (int)TabSize };
        using var cancellation = new CancellationTokenSource();
        _cancellation = cancellation;
        IsBusy = true;
        var progress = new ArchiveProgress();
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        timer.Tick += (_, _) => Status = _localizer.FormatOrDefault("Conversion.Archive.Progress", "Completed {0}; failed {1}. Completed files survive cancellation.", progress.Completed, progress.Failed);
        Status = _localizer.GetOrDefault("Conversion.Archive.Running", "Running archive. Completed files are checkpointed individually.");
        timer.Start();
        try {
            OfficePdfArchiveResult result = await OfficePdfArchiveWorkflow.RunAsync(request, progress, cancellation.Token, _guard).ConfigureAwait(true);
            Status = _localizer.FormatOrDefault(result.Cancelled ? "Conversion.Archive.CancelledResult" : "Conversion.Archive.FinishedResult",
                result.Cancelled ? "Cancelled: {0} completed, {1} reused, {2} failed. Source and output hashes are verified on rerun." : "Finished: {0} completed, {1} reused, {2} failed. Source and output hashes are verified on rerun.",
                result.Completed, result.Reused, result.Failed);
            if (progress.FirstFailure is { } failure) Status += " " + failure;
        } catch (OperationCanceledException) { Status = _localizer.GetOrDefault("Conversion.Archive.Cancelled", "Archive cancelled. Completed files remain available for resume."); }
        catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) { Status = error.Message; }
        finally { timer.Stop(); _cancellation = null; IsBusy = false; }
    }
    [RelayCommand(CanExecute = nameof(IsBusy))]
    private void Cancel() => _cancellation?.Cancel();
    /// <inheritdoc />
    public void Dispose() => _cancellation?.Cancel();
    private sealed class ArchiveProgress : IProgress<OfficePdfArchiveItemResult> {
        private long _completed, _failed;
        private string? _firstFailure;
        public string? FirstFailure => Volatile.Read(ref _firstFailure);
        public long Completed => Interlocked.Read(ref _completed);
        public long Failed => Interlocked.Read(ref _failed);
        public void Report(OfficePdfArchiveItemResult item) {
            if (item.Status == OfficeWorkflowStatus.Completed) Interlocked.Increment(ref _completed);
            else if (item.Status == OfficeWorkflowStatus.Failed) {
                Interlocked.Increment(ref _failed);
                string message = Path.GetFileName(item.InputPath) + ": " + item.Summary;
                Interlocked.CompareExchange(ref _firstFailure, message.Length > 1024 ? message[..1024] : message, null);
            }
        }
    }
}
