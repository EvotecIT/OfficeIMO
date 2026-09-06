using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class PdfAssemblySourceViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;

    public PdfAssemblySourceViewModel(string path) : this(path, null) { }

    internal PdfAssemblySourceViewModel(string path, IStudioLocalizer? localizer, string? name = null) {
        Path = OfficeStorageIdentity.Normalize(path);
        _name = name;
        _localizer = localizer ?? StudioLocalization.Current;
    }
    public string Path { get; }
    private readonly string? _name;
    public string Name => _name ?? (Directory.Exists(Path) ? new DirectoryInfo(Path).Name : System.IO.Path.GetFileName(Path));
    public string Kind => Directory.Exists(Path)
        ? _localizer.GetOrDefault("Assembly.Source.Folder", "Folder")
        : System.IO.Path.GetExtension(Name).TrimStart('.').ToUpperInvariant();
}

public sealed partial class PdfAssemblyViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickFolder;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputPdf;
    private readonly IOfficeOutputWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private readonly IOfficeWorkflowPublicationGuard? _publicationGuard;
    private readonly StudioJobHistory? _jobHistory;
    private readonly StudioStorageAccess? _storage;
    private CancellationTokenSource? _cancellation;

    public PdfAssemblyViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickFolder,
        Func<CancellationToken, Task<string?>> pickOutputPdf,
        IOfficeOutputWorkflowRunner? runner = null) : this(pickFiles, pickFolder, pickOutputPdf, runner, null) { }

    internal PdfAssemblyViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickFolder,
        Func<CancellationToken, Task<string?>> pickOutputPdf,
        IOfficeOutputWorkflowRunner? runner,
        IStudioLocalizer? localizer = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null,
        StudioJobHistory? jobHistory = null,
        StudioStorageAccess? storage = null) {
        _pickFiles = pickFiles;
        _pickFolder = pickFolder;
        _pickOutputPdf = pickOutputPdf;
        _runner = runner ?? new OfficeWorkflowRunner();
        _publicationGuard = publicationGuard;
        _jobHistory = jobHistory;
        _storage = storage;
        _localizer = localizer ?? StudioLocalization.Current;
        Status = T("Status.Ready", "Add documents, images, folders, or ZIPs in the order you want.");
        Summary = T("Summary.Empty", "No assembly run yet");
    }

    public ObservableCollection<PdfAssemblySourceViewModel> Sources { get; } = new();

    [ObservableProperty]
    private PdfAssemblySourceViewModel? _selectedSource;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private string _outputPath = string.Empty;

    [ObservableProperty]
    private bool _includeSubdirectories = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanCancel))]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private string _summary = string.Empty;

    [ObservableProperty]
    private string? _publishedPath;

    public bool HasSources => Sources.Count > 0;
    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(PublishedPath);
    public string SourceSummary => Sources.Count == 0
        ? T("Sources.Empty", "No sources")
        : _localizer.FormatOrDefault("Assembly.Sources.Count", "{0:N0} {1}", Sources.Count, Sources.Count == 1 ? T("Sources.One", "source") : T("Sources.Many", "sources"));
    private bool CanRun => !IsBusy && HasSources && !string.IsNullOrWhiteSpace(OutputPath);

    internal void UseDocument(string? path) {
        if (!string.IsNullOrWhiteSpace(path)) AddSources([path]);
    }

    [RelayCommand]
    private async Task AddFilesAsync(CancellationToken cancellationToken) {
        IReadOnlyList<string> paths = await _pickFiles(cancellationToken).ConfigureAwait(true);
        AddSources(paths);
    }

    [RelayCommand]
    private async Task AddFolderAsync(CancellationToken cancellationToken) {
        string? path = await _pickFolder(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) AddSources([path]);
    }

    [RelayCommand]
    private async Task ChooseOutputAsync(CancellationToken cancellationToken) {
        string? path = await _pickOutputPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) OutputPath = path;
    }

    [RelayCommand]
    private void RemoveSelected() {
        if (IsBusy || SelectedSource is null) return;
        int index = Sources.IndexOf(SelectedSource);
        Sources.Remove(SelectedSource);
        SelectedSource = Sources.Count == 0 ? null : Sources[Math.Min(index, Sources.Count - 1)];
        NotifySourcesChanged();
    }

    [RelayCommand]
    private void MoveUp() => MoveSelected(-1);

    [RelayCommand]
    private void MoveDown() => MoveSelected(1);

    [RelayCommand]
    private void Clear() {
        if (IsBusy) return;
        Sources.Clear();
        SelectedSource = null;
        NotifySourcesChanged();
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync() {
        _cancellation?.Dispose();
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        IsBusy = true;
        ProgressFraction = 0D;
        PublishedPath = null;
        OnPropertyChanged(nameof(HasOutput));

        StudioJobRecord? job = null;
        bool ownerStarted = false;

        try {
            var request = new PdfAssemblyRequest {
                Sources = Sources.Select(static source => source.Path).ToArray(),
                SourceStreams = Sources.Select(source => (source.Path, Access: _storage?.CreateWorkflowInput(source.Path)))
                    .Where(source => source.Access is not null).ToDictionary(source => source.Path, source => source.Access!, StringComparer.Ordinal),
                OutputPath = OutputPath,
                PublicationGuard = _publicationGuard,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename,
                Options = new PdfAssemblyOptions { IncludeSubdirectories = IncludeSubdirectories }
            };
            job = _jobHistory?.Start(T("Job.Title", "PDF assembly"), string.Join(Environment.NewLine, request.Sources), request.OutputPath, operation.Cancel);
            using IDisposable? execution = _jobHistory is null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                if (!IsBusy || !ReferenceEquals(_cancellation, operation)) return;
                ProgressFraction = update.Fraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
                job?.Report(update);
            });
            ownerStarted = true;
            PdfAssemblyResult result = await _runner.AssemblePdfAsync(request, progress, operation.Token).ConfigureAwait(true);
            job?.Complete(result.Status, result.OutputPath, result.Summary);
            Summary = result.Summary;
            Status = result.Status switch {
                OfficeWorkflowStatus.Completed => T("Status.Completed", "Assembled PDF ready"),
                OfficeWorkflowStatus.Cancelled => T("Status.Cancelled", "Assembly cancelled"),
                _ => _localizer.GetOrDefault("Assembly.Status.Failed", result.Summary)
            };
            PublishedPath = result.OutputPath;
            ProgressFraction = result.Status == OfficeWorkflowStatus.Completed ? 1D : ProgressFraction;
            OnPropertyChanged(nameof(HasOutput));
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = ownerStarted ? _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output")
                : _localizer.GetOrDefault("Workflow.Status.Cancelled", "Cancelled");
            if (ownerStarted) job?.Unconfirmed(Status);
            else job?.Complete(OfficeWorkflowStatus.Cancelled, null, Status);
        } catch (Exception exception) {
            Status = exception.Message;
            job?.Unconfirmed(exception.Message);
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    private void AddSources(IEnumerable<string> paths) {
        if (IsBusy) return;
        var existing = new HashSet<string>(StringComparer.Ordinal);
        try {
            foreach (var source in Sources) existing.Add(InputIdentity(source.Path));
        } catch (Exception exception) when (IsIdentityFailure(exception)) {
            Status = T("Sources.Unavailable", "An assembly source could not be inspected. Restore or remove it before adding more sources.");
            return;
        }
        int skipped = 0;
        foreach (string path in paths.Where(static path => !string.IsNullOrWhiteSpace(path))) {
            try {
                string fullPath = OfficeStorageIdentity.Normalize(path);
                if (!existing.Add(InputIdentity(fullPath))) continue;
                Sources.Add(new PdfAssemblySourceViewModel(fullPath, _localizer, _storage?.Describe(fullPath).Name));
            } catch (Exception exception) when (IsIdentityFailure(exception)) { skipped++; }
        }
        SelectedSource ??= Sources.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(OutputPath) && Sources.Count > 0 && _storage?.UsesProviderPublication(Sources[0].Path) != true) {
            string first = Sources[0].Path;
            string directory = Directory.Exists(first)
                ? Directory.GetParent(first)?.FullName ?? first
                : System.IO.Path.GetDirectoryName(first)!;
            OutputPath = System.IO.Path.Combine(directory, "assembled.pdf");
        }
        NotifySourcesChanged();
        if (skipped > 0) Status = _localizer.FormatOrDefault("Assembly.Sources.Skipped", "Skipped {0:N0} source(s) that could not be inspected.", skipped);
    }

    private static bool IsIdentityFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException;

    private string InputIdentity(string location) => _storage?.UsesProviderPublication(location) == true
        ? OfficeStorageIdentity.Normalize(location) : OfficePathIdentity.GetPathIdentityKey(location);

    private void MoveSelected(int offset) {
        if (IsBusy || SelectedSource is null) return;
        int oldIndex = Sources.IndexOf(SelectedSource);
        int newIndex = oldIndex + offset;
        if (newIndex < 0 || newIndex >= Sources.Count) return;
        Sources.Move(oldIndex, newIndex);
    }

    private void NotifySourcesChanged() {
        OnPropertyChanged(nameof(HasSources));
        OnPropertyChanged(nameof(SourceSummary));
        RunCommand.NotifyCanExecuteChanged();
    }

    public void Dispose() => _cancellation?.Cancel();

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Assembly." + suffix, fallback);
}
