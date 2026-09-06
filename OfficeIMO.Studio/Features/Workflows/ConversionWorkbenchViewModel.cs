using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private readonly IOfficeWorkflowPublicationGuard? _publicationGuard;
    private readonly StudioJobHistory? _jobHistory;
    private readonly StudioStorageAccess? _storage;
    private CancellationTokenSource? _cancellation;

    public ConversionWorkbenchViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner = null) : this(pickFiles, pickOutputFolder, runner, null) { }

    internal ConversionWorkbenchViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner,
        IStudioLocalizer? localizer = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null,
        StudioJobHistory? jobHistory = null,
        StudioStorageAccess? storage = null) {
        _pickFiles = pickFiles;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        _publicationGuard = publicationGuard;
        _jobHistory = jobHistory;
        _storage = storage;
        _localizer = localizer ?? StudioLocalization.Current;
        Routes = OfficeWorkflowCatalog.Routes.Select(route => new ConversionRouteChoice(route, _localizer)).ToArray();
        Profiles = [
            new(OfficeWorkflowOutputProfile.Faithful, T("Profile.Faithful.Label", "Faithful"), T("Profile.Faithful.Description", "Preserve authored content and visual features where the format owner supports them.")),
            new(OfficeWorkflowOutputProfile.Lightweight, T("Profile.Lightweight.Label", "Lightweight"), T("Profile.Lightweight.Description", "Prefer smaller, simpler output while retaining useful structure.")),
            new(OfficeWorkflowOutputProfile.PrintReady, T("Profile.PrintReady.Label", "Print ready"), T("Profile.PrintReady.Description", "Prefer pagination and document setup intended for printing.")),
            new(OfficeWorkflowOutputProfile.TextOnly, T("Profile.TextOnly.Label", "Text focused"), T("Profile.TextOnly.Description", "Prefer text and tables over decorative visual content."))
        ];
        ConflictPolicies = [
            new(OfficeWorkflowConflictPolicy.Rename, T("Conflict.Rename.Label", "Create numbered copy"), T("Conflict.Rename.Description", "Keep the existing file and add a numbered suffix.")),
            new(OfficeWorkflowConflictPolicy.Fail, T("Conflict.Fail.Label", "Stop that job"), T("Conflict.Fail.Description", "Report the collision without changing the existing file.")),
            new(OfficeWorkflowConflictPolicy.Replace, T("Conflict.Replace.Label", "Replace after validation"), T("Conflict.Replace.Description", "Replace only after the new artifact passes reopen validation."))
        ];
        SelectedRoute = Routes.First();
        SelectedProfile = Profiles[0];
        SelectedConflict = ConflictPolicies[0];
        Status = T("Status.Ready", "Choose a route, then add one or more matching files.");
    }

    public IReadOnlyList<ConversionRouteChoice> Routes { get; }

    public IReadOnlyList<WorkflowProfileChoice> Profiles { get; }

    public IReadOnlyList<WorkflowConflictChoice> ConflictPolicies { get; }

    public ObservableCollection<ConversionJobViewModel> Jobs { get; } = new();

    [ObservableProperty]
    private ConversionRouteChoice _selectedRoute = null!;

    [ObservableProperty]
    private WorkflowProfileChoice _selectedProfile = null!;

    [ObservableProperty]
    private WorkflowConflictChoice _selectedConflict = null!;

    [ObservableProperty]
    private ConversionJobViewModel? _selectedJob;

    [ObservableProperty]
    private string _outputFolder = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanRun))]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    private string _status = string.Empty;

    public bool HasJobs => Jobs.Count > 0;
    public bool CanRun => !IsBusy && Jobs.Any(job => job.State == ConversionJobState.Queued);
    public bool CanRetryFailed => !IsBusy && Jobs.Any(job => job.CanRetry);
    public bool CanCancel => IsBusy;
    public bool CanEditQueue => !IsBusy;
    public string QueueSummary => Jobs.Count == 0
        ? T("Queue.Empty", "No jobs")
        : _localizer.FormatOrDefault("Conversion.Queue.Count", "{0:N0} {1}", Jobs.Count, Jobs.Count == 1 ? T("Queue.Job", "job") : T("Queue.Jobs", "jobs"));

    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanCancel));
        OnPropertyChanged(nameof(CanEditQueue));
        OnPropertyChanged(nameof(CanRetryFailed));
        RunQueueCommand.NotifyCanExecuteChanged();
        RetryFailedCommand.NotifyCanExecuteChanged();
        AddFilesCommand.NotifyCanExecuteChanged();
        RemoveSelectedCommand.NotifyCanExecuteChanged();
        ClearQueueCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanEditQueue))]
    private async Task AddFilesAsync(CancellationToken cancellationToken) {
        IReadOnlyList<string> paths = await _pickFiles(cancellationToken).ConfigureAwait(true);
        if (IsBusy) return;
        int added = 0;
        int skipped = 0;
        int skippedForLimit = 0;
        var identities = new HashSet<string>(StringComparer.Ordinal);
        try {
            foreach (ConversionJobViewModel job in Jobs.Where(job => job.Route.Route.Id == SelectedRoute.Route.Id)) {
                identities.Add(InputIdentity(job.InputPath));
            }
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            Status = _localizer.FormatOrDefault("Conversion.Add.IdentityFailed", "A queued input could not be inspected. Remove or restore it before adding files: {0}", exception.Message);
            return;
        }
        foreach (string path in paths) {
            string fileName = _storage?.Describe(path).Name ?? Path.GetFileName(path);
            string extension = Path.GetExtension(fileName);
            bool accepts = SelectedRoute.Route.SourceExtensions.Any(item =>
                string.Equals(NormalizeExtension(item), extension, StringComparison.OrdinalIgnoreCase));
            if (!accepts) {
                skipped++;
                continue;
            }
            string fullPath;
            try {
                fullPath = OfficeStorageIdentity.Normalize(path);
                if (!identities.Add(InputIdentity(fullPath))) {
                    skipped++;
                    continue;
                }
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
                skipped++;
                continue;
            }
            if (Jobs.Count >= OfficeWorkflowRunner.MaximumBatchRequestCount) {
                skippedForLimit++;
                continue;
            }
            var job = new ConversionJobViewModel(fullPath, SelectedRoute, _localizer, fileName);
            Jobs.Add(job);
            SelectedJob ??= job;
            added++;
        }
        NotifyQueueChanged();
        Status = skippedForLimit > 0
            ? _localizer.FormatOrDefault("Conversion.Queue.Limit", "The queue is limited to {0:N0} jobs; {1:N0} additional file(s) were not added.", OfficeWorkflowRunner.MaximumBatchRequestCount, skippedForLimit)
            : added == 0
            ? _localizer.FormatOrDefault("Conversion.Add.NoNewFiles", "No files matched {0} that could be added; files already queued for this route were skipped.", SelectedRoute.Route.Source)
            : skipped == 0
                ? _localizer.FormatOrDefault("Conversion.Add.Success", "Added {0:N0} {1}.", added, added == 1 ? T("Queue.File", "file") : T("Queue.Files", "files"))
                : _localizer.FormatOrDefault("Conversion.Add.SkippedInputs", "Added {0:N0}; skipped {1:N0} duplicate, unsupported, or uninspectable file(s).", added, skipped);
    }

    [RelayCommand]
    private async Task ChooseOutputFolderAsync(CancellationToken cancellationToken) {
        string? path = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) OutputFolder = path;
    }

    [RelayCommand(CanExecute = nameof(CanEditQueue))]
    private void RemoveSelected() {
        if (IsBusy || SelectedJob is null) return;
        int index = Jobs.IndexOf(SelectedJob);
        Jobs.Remove(SelectedJob);
        SelectedJob = Jobs.Count == 0 ? null : Jobs[Math.Min(index, Jobs.Count - 1)];
        NotifyQueueChanged();
    }

    [RelayCommand(CanExecute = nameof(CanEditQueue))]
    private void ClearQueue() {
        if (IsBusy) return;
        Jobs.Clear();
        SelectedJob = null;
        ProgressFraction = 0D;
        Status = T("Queue.Cleared", "Queue cleared.");
        NotifyQueueChanged();
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task RunQueueAsync() => RunJobsAsync(Jobs.Where(job => job.State == ConversionJobState.Queued).ToArray());

    [RelayCommand(CanExecute = nameof(CanRetryFailed))]
    private Task RetryFailedAsync() => RunJobsAsync(Jobs.Where(job => job.CanRetry).ToArray());

    private async Task RunJobsAsync(ConversionJobViewModel[] candidates) {
        if (IsBusy || candidates.Length == 0) return;
        if (string.IsNullOrWhiteSpace(OutputFolder) && candidates.Any(job => _storage?.UsesProviderPublication(job.InputPath) == true)) {
            Status = T("Output.ProviderFolderRequired", "Choose an output folder before converting provider documents.");
            return;
        }
        _cancellation?.Dispose();
        var operationCancellation = new CancellationTokenSource();
        _cancellation = operationCancellation;
        IsBusy = true;
        ProgressFraction = 0D;
        foreach (ConversionJobViewModel job in candidates) job.PrepareAttempt();
        var history = new Dictionary<string, StudioJobRecord>(StringComparer.Ordinal);
        bool ownerStarted = false;

        try {
            OfficeWorkflowRequest[] requests = candidates.Select(CreateRequest).ToArray();
            if (_jobHistory is not null) {
                foreach (OfficeWorkflowRequest request in requests) {
                    history.Add(request.Id, _jobHistory.Start(T("Job.Title", "Conversion"), request.InputPath,
                        request.OutputPath, operationCancellation.Cancel, batch: candidates.Length > 1));
                }
            }
            using IDisposable? execution = _jobHistory is null ? null : await _jobHistory.EnterAsync(operationCancellation.Token).ConfigureAwait(true);
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                if (!ReferenceEquals(_cancellation, operationCancellation) || !IsBusy) return;
                ProgressFraction = update.OverallFraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
                candidates.FirstOrDefault(item => item.Id == update.RequestId)?.ReportProgress(update);
                if (history.TryGetValue(update.RequestId, out StudioJobRecord? entry)) entry.Report(update);
            });
            ownerStarted = true;
            IReadOnlyList<OfficeWorkflowResult> results = await _runner
                .RunBatchAsync(requests, progress, operationCancellation.Token)
                .ConfigureAwait(true);
            var received = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeWorkflowResult result in results) {
                ConversionJobViewModel? job = candidates.FirstOrDefault(item => item.Id == result.RequestId);
                if (job is null || !received.Add(result.RequestId)) {
                    throw new InvalidOperationException("The workflow runner returned an unexpected job result.");
                }
                job.Apply(result);
                if (history.TryGetValue(result.RequestId, out StudioJobRecord? entry)) entry.Complete(result.Status, result.OutputPath, result.Summary);
            }
            foreach (ConversionJobViewModel job in candidates.Where(job => !received.Contains(job.Id))) {
                job.EndWithoutResult(operationCancellation.IsCancellationRequested,
                    operationCancellation.IsCancellationRequested
                        ? T("Job.NotStarted", "Cancelled before this job started; completed outputs were retained.")
                        : T("Job.NoResult", "No result was returned. Check the output folder before starting another attempt."));
                if (history.TryGetValue(job.Id, out StudioJobRecord? entry)) {
                    if (operationCancellation.IsCancellationRequested) entry.Complete(OfficeWorkflowStatus.Cancelled, null, job.Summary!);
                    else entry.Unconfirmed(job.Summary!);
                }
            }
            Status = _localizer.FormatOrDefault(
                "Conversion.Queue.Outcomes",
                "{0:N0} completed · {1:N0} failed · {2:N0} cancelled · {3:N0} need checking",
                Jobs.Count(job => job.State == ConversionJobState.Completed),
                Jobs.Count(job => job.State == ConversionJobState.Failed),
                Jobs.Count(job => job.State == ConversionJobState.Cancelled),
                Jobs.Count(job => job.State == ConversionJobState.Unconfirmed));
            ProgressFraction = 1D;
        } catch (Exception exception) {
            string message = T("Job.UnconfirmedOutput", "The runner stopped without a result. Check the output folder before starting another attempt.");
            bool cancelledBeforeStart = !ownerStarted && exception is OperationCanceledException && operationCancellation.IsCancellationRequested;
            if (cancelledBeforeStart) message = T("Job.NotStarted", "Cancelled before this job started; completed outputs were retained.");
            foreach (ConversionJobViewModel job in candidates.Where(job => job.State is ConversionJobState.Queued or ConversionJobState.Running)) {
                job.EndWithoutResult(cancelledBeforeStart, message);
            }
            foreach (StudioJobRecord entry in history.Values.Where(entry => entry.IsActive)) {
                if (cancelledBeforeStart) entry.Complete(OfficeWorkflowStatus.Cancelled, null, message);
                else entry.Unconfirmed(message);
            }
            Status = _localizer.FormatOrDefault("Conversion.Queue.Failed", "The conversion queue could not finish: {0}", exception.Message);
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operationCancellation)) _cancellation = null;
            operationCancellation.Dispose();
            NotifyQueueChanged();
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    public void Dispose() {
        _cancellation?.Cancel();
    }

    private OfficeWorkflowRequest CreateRequest(ConversionJobViewModel job) {
        if (string.IsNullOrWhiteSpace(OutputFolder) && _storage?.UsesProviderPublication(job.InputPath) == true) {
            throw new InvalidOperationException(T("Output.ProviderFolderRequired", "Choose an output folder before converting provider documents."));
        }
        string directory = string.IsNullOrWhiteSpace(OutputFolder)
            ? Path.GetDirectoryName(job.InputPath)!
            : Path.GetFullPath(OutputFolder);
        string outputPath = Path.Combine(
            directory,
            Path.GetFileNameWithoutExtension(job.FileName) + NormalizeExtension(job.Route.Route.TargetExtension));
        return new OfficeWorkflowRequest {
            Id = job.Id,
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = job.InputPath,
            InputStream = _storage?.CreateWorkflowInput(job.InputPath),
            OutputPath = outputPath,
            ConversionRouteId = job.Route.Route.Id,
            OutputProfile = SelectedProfile.Value,
            PublicationGuard = _publicationGuard,
            ConflictPolicy = SelectedConflict.Value
        };
    }

    private void NotifyQueueChanged() {
        OnPropertyChanged(nameof(HasJobs));
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(CanRetryFailed));
        OnPropertyChanged(nameof(QueueSummary));
        RunQueueCommand.NotifyCanExecuteChanged();
        RetryFailedCommand.NotifyCanExecuteChanged();
    }

    private static string NormalizeExtension(string extension) => extension.StartsWith('.') ? extension : "." + extension;

    private string InputIdentity(string location) => _storage?.UsesProviderPublication(location) == true
        ? OfficeStorageIdentity.Normalize(location) : OfficePathIdentity.GetPathIdentityKey(location);

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Conversion." + suffix, fallback);
}
