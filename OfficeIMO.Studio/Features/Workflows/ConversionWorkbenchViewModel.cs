using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeWorkflowRunner _runner;
    private CancellationTokenSource? _cancellation;

    public ConversionWorkbenchViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner = null) {
        _pickFiles = pickFiles;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        Routes = OfficeWorkflowCatalog.Routes.Select(route => new ConversionRouteChoice(route)).ToArray();
        SelectedRoute = Routes.First();
        SelectedProfile = Profiles[0];
        SelectedConflict = ConflictPolicies[0];
    }

    public IReadOnlyList<ConversionRouteChoice> Routes { get; }

    public IReadOnlyList<WorkflowProfileChoice> Profiles { get; } = [
        new(OfficeWorkflowOutputProfile.Faithful, "Faithful", "Preserve authored content and visual features where the format owner supports them."),
        new(OfficeWorkflowOutputProfile.Lightweight, "Lightweight", "Prefer smaller, simpler output while retaining useful structure."),
        new(OfficeWorkflowOutputProfile.PrintReady, "Print ready", "Prefer pagination and document setup intended for printing."),
        new(OfficeWorkflowOutputProfile.TextOnly, "Text focused", "Prefer text and tables over decorative visual content.")
    ];

    public IReadOnlyList<WorkflowConflictChoice> ConflictPolicies { get; } = [
        new(OfficeWorkflowConflictPolicy.Rename, "Create numbered copy", "Keep the existing file and add a numbered suffix."),
        new(OfficeWorkflowConflictPolicy.Fail, "Stop that job", "Report the collision without changing the existing file."),
        new(OfficeWorkflowConflictPolicy.Replace, "Replace after validation", "Replace only after the new artifact passes reopen validation.")
    ];

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
    private string _status = "Choose a route, then add one or more matching files.";

    public bool HasJobs => Jobs.Count > 0;
    public bool CanRun => HasJobs && !IsBusy;
    public bool CanCancel => IsBusy;
    public bool CanEditQueue => !IsBusy;
    public string QueueSummary => Jobs.Count == 0 ? "No jobs" : $"{Jobs.Count:N0} {(Jobs.Count == 1 ? "job" : "jobs")}";

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
        foreach (string path in paths.Distinct(StringComparer.OrdinalIgnoreCase)) {
            string extension = Path.GetExtension(path);
            bool accepts = SelectedRoute.Route.SourceExtensions.Any(item =>
                string.Equals(NormalizeExtension(item), extension, StringComparison.OrdinalIgnoreCase));
            if (!accepts) {
                skipped++;
                continue;
            }
            var job = new ConversionJobViewModel(Path.GetFullPath(path), SelectedRoute);
            Jobs.Add(job);
            SelectedJob ??= job;
            added++;
        }
        NotifyQueueChanged();
        Status = added == 0
            ? $"No files matched {SelectedRoute.Route.Source}."
            : skipped == 0
                ? $"Added {added:N0} {(added == 1 ? "file" : "files")}."
                : $"Added {added:N0}; skipped {skipped:N0} file(s) that do not match this route.";
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
        Status = "Queue cleared.";
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
            job.Status = "Queued";
            job.ProgressFraction = 0D;
        }

        try {
            OfficeWorkflowRequest[] requests = Jobs.Select(CreateRequest).ToArray();
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.OverallFraction;
                Status = update.Message;
                ConversionJobViewModel? job = Jobs.FirstOrDefault(item => item.Id == update.RequestId);
                if (job is not null) {
                    job.Status = update.Stage == "complete" ? "Completed" : "Running · " + update.Stage.Replace('-', ' ');
                    job.ProgressFraction = update.Fraction;
                }
            });
            IReadOnlyList<OfficeWorkflowResult> results = await _runner
                .RunBatchAsync(requests, progress, operationCancellation.Token)
                .ConfigureAwait(true);
            foreach (OfficeWorkflowResult result in results) {
                Jobs.First(job => job.Id == result.RequestId).Apply(result);
            }
            foreach (ConversionJobViewModel job in Jobs.Where(job => job.Status == "Queued")) {
                job.Status = "Cancelled";
            }
            int completed = results.Count(result => result.Succeeded);
            int failed = results.Count(result => result.Status == OfficeWorkflowStatus.Failed);
            int cancelled = Jobs.Count - completed - failed;
            Status = $"Queue finished · {completed:N0} completed · {failed:N0} failed · {cancelled:N0} cancelled";
            ProgressFraction = results.Count == Jobs.Count ? 1D : ProgressFraction;
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
}
