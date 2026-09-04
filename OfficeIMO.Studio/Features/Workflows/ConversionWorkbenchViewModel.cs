using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _cancellation;

    public ConversionWorkbenchViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner = null) : this(pickFiles, pickOutputFolder, runner, null) { }

    internal ConversionWorkbenchViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner,
        IStudioLocalizer? localizer = null) {
        _pickFiles = pickFiles;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
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
    public bool CanRun => HasJobs && !IsBusy;
    public bool CanCancel => IsBusy;
    public bool CanEditQueue => !IsBusy;
    public string QueueSummary => Jobs.Count == 0
        ? T("Queue.Empty", "No jobs")
        : _localizer.FormatOrDefault("Conversion.Queue.Count", "{0:N0} {1}", Jobs.Count, Jobs.Count == 1 ? T("Queue.Job", "job") : T("Queue.Jobs", "jobs"));

    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanCancel));
        OnPropertyChanged(nameof(CanEditQueue));
        RunQueueCommand.NotifyCanExecuteChanged();
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
        foreach (string path in paths.Distinct(StringComparer.OrdinalIgnoreCase)) {
            string extension = Path.GetExtension(path);
            bool accepts = SelectedRoute.Route.SourceExtensions.Any(item =>
                string.Equals(NormalizeExtension(item), extension, StringComparison.OrdinalIgnoreCase));
            if (!accepts) {
                skipped++;
                continue;
            }
            if (Jobs.Count >= OfficeWorkflowRunner.MaximumBatchRequestCount) {
                skippedForLimit++;
                continue;
            }
            var job = new ConversionJobViewModel(Path.GetFullPath(path), SelectedRoute, _localizer);
            Jobs.Add(job);
            SelectedJob ??= job;
            added++;
        }
        NotifyQueueChanged();
        Status = skippedForLimit > 0
            ? _localizer.FormatOrDefault("Conversion.Queue.Limit", "The queue is limited to {0:N0} jobs; {1:N0} additional file(s) were not added.", OfficeWorkflowRunner.MaximumBatchRequestCount, skippedForLimit)
            : added == 0
            ? _localizer.FormatOrDefault("Conversion.Add.None", "No files matched {0}.", SelectedRoute.Route.Source)
            : skipped == 0
                ? _localizer.FormatOrDefault("Conversion.Add.Success", "Added {0:N0} {1}.", added, added == 1 ? T("Queue.File", "file") : T("Queue.Files", "files"))
                : _localizer.FormatOrDefault("Conversion.Add.Partial", "Added {0:N0}; skipped {1:N0} file(s) that do not match this route.", added, skipped);
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
    private async Task RunQueueAsync() {
        _cancellation?.Dispose();
        var operationCancellation = new CancellationTokenSource();
        _cancellation = operationCancellation;
        IsBusy = true;
        ProgressFraction = 0D;
        foreach (ConversionJobViewModel job in Jobs) {
            job.Status = T("Job.Queued", "Queued");
            job.ProgressFraction = 0D;
        }

        try {
            OfficeWorkflowRequest[] requests = Jobs.Select(CreateRequest).ToArray();
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.OverallFraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
                ConversionJobViewModel? job = Jobs.FirstOrDefault(item => item.Id == update.RequestId);
                if (job is not null) {
                    job.Status = update.Stage == "complete"
                        ? T("Job.Completed", "Completed")
                        : _localizer.FormatOrDefault("Conversion.Job.Running", "Running · {0}", update.Stage.Replace('-', ' '));
                    job.ProgressFraction = update.Fraction;
                }
            });
            IReadOnlyList<OfficeWorkflowResult> results = await _runner
                .RunBatchAsync(requests, progress, operationCancellation.Token)
                .ConfigureAwait(true);
            foreach (OfficeWorkflowResult result in results) {
                Jobs.First(job => job.Id == result.RequestId).Apply(result);
            }
            foreach (ConversionJobViewModel job in Jobs.Where(job => job.ProgressFraction == 0D)) {
                job.Status = T("Job.Cancelled", "Cancelled");
            }
            int completed = results.Count(result => result.Succeeded);
            int failed = results.Count(result => result.Status == OfficeWorkflowStatus.Failed);
            int cancelled = Jobs.Count - completed - failed;
            Status = _localizer.FormatOrDefault(
                "Conversion.Queue.Finished",
                "Queue finished · {0:N0} completed · {1:N0} failed · {2:N0} cancelled",
                completed,
                failed,
                cancelled);
            ProgressFraction = results.Count == Jobs.Count ? 1D : ProgressFraction;
        } catch (OperationCanceledException) when (operationCancellation.IsCancellationRequested) {
            Status = T("Queue.Cancelled", "Queue cancelled.");
        } catch (Exception exception) {
            Status = _localizer.FormatOrDefault("Conversion.Queue.Failed", "The conversion queue could not finish: {0}", exception.Message);
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operationCancellation)) _cancellation = null;
            operationCancellation.Dispose();
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    public void Dispose() {
        _cancellation?.Cancel();
    }

    private OfficeWorkflowRequest CreateRequest(ConversionJobViewModel job) {
        string directory = string.IsNullOrWhiteSpace(OutputFolder)
            ? Path.GetDirectoryName(job.InputPath)!
            : Path.GetFullPath(OutputFolder);
        string outputPath = Path.Combine(
            directory,
            Path.GetFileNameWithoutExtension(job.InputPath) + NormalizeExtension(job.Route.Route.TargetExtension));
        return new OfficeWorkflowRequest {
            Id = job.Id,
            Operation = OfficeWorkflowOperation.Convert,
            InputPath = job.InputPath,
            OutputPath = outputPath,
            ConversionRouteId = job.Route.Route.Id,
            OutputProfile = SelectedProfile.Value,
            ConflictPolicy = SelectedConflict.Value
        };
    }

    private void NotifyQueueChanged() {
        OnPropertyChanged(nameof(HasJobs));
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(QueueSummary));
        RunQueueCommand.NotifyCanExecuteChanged();
    }

    private static string NormalizeExtension(string extension) => extension.StartsWith('.') ? extension : "." + extension;

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Conversion." + suffix, fallback);
}
