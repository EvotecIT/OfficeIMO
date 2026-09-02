using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class PdfAssemblySourceViewModel : ObservableObject {
    public PdfAssemblySourceViewModel(string path) => Path = System.IO.Path.GetFullPath(path);
    public string Path { get; }
    public string Name => Directory.Exists(Path) ? new DirectoryInfo(Path).Name : System.IO.Path.GetFileName(Path);
    public string Kind => Directory.Exists(Path) ? "Folder" : System.IO.Path.GetExtension(Path).TrimStart('.').ToUpperInvariant();
}

public sealed partial class PdfAssemblyViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickFolder;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputPdf;
    private readonly IOfficeOutputWorkflowRunner _runner;
    private CancellationTokenSource? _cancellation;

    public PdfAssemblyViewModel(
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickFolder,
        Func<CancellationToken, Task<string?>> pickOutputPdf,
        IOfficeOutputWorkflowRunner? runner = null) {
        _pickFiles = pickFiles;
        _pickFolder = pickFolder;
        _pickOutputPdf = pickOutputPdf;
        _runner = runner ?? new OfficeWorkflowRunner();
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
    private string _status = "Add documents, images, folders, or ZIPs in the order you want.";

    [ObservableProperty]
    private string _summary = "No assembly run yet";

    [ObservableProperty]
    private string? _publishedPath;

    public bool HasSources => Sources.Count > 0;
    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(PublishedPath);
    public string SourceSummary => Sources.Count == 0 ? "No sources" : $"{Sources.Count:N0} {(Sources.Count == 1 ? "source" : "sources")}";
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

        try {
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.Fraction;
                Status = update.Message;
            });
            PdfAssemblyResult result = await _runner.AssemblePdfAsync(new PdfAssemblyRequest {
                Sources = Sources.Select(static source => source.Path).ToArray(),
                OutputPath = OutputPath,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename,
                Options = new PdfAssemblyOptions { IncludeSubdirectories = IncludeSubdirectories }
            }, progress, operation.Token).ConfigureAwait(true);
            Summary = result.Summary;
            Status = result.Status switch {
                OfficeWorkflowStatus.Completed => "Assembled PDF ready",
                OfficeWorkflowStatus.Cancelled => "Assembly cancelled",
                _ => result.Summary
            };
            PublishedPath = result.OutputPath;
            ProgressFraction = result.Status == OfficeWorkflowStatus.Completed ? 1D : ProgressFraction;
            OnPropertyChanged(nameof(HasOutput));
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    private void AddSources(IEnumerable<string> paths) {
        var existing = Sources.Select(static source => source.Path).ToList();
        foreach (string path in paths.Where(static path => !string.IsNullOrWhiteSpace(path))) {
            string fullPath = System.IO.Path.GetFullPath(path);
            if (existing.Any(candidate => AreEquivalentPaths(candidate, fullPath))) continue;
            existing.Add(fullPath);
            Sources.Add(new PdfAssemblySourceViewModel(fullPath));
        }
        SelectedSource ??= Sources.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(OutputPath) && Sources.Count > 0) {
            string first = Sources[0].Path;
            string directory = Directory.Exists(first)
                ? Directory.GetParent(first)?.FullName ?? first
                : System.IO.Path.GetDirectoryName(first)!;
            OutputPath = System.IO.Path.Combine(directory, "assembled.pdf");
        }
        NotifySourcesChanged();
    }

    private static bool AreEquivalentPaths(string left, string right) {
        string normalizedLeft = System.IO.Path.TrimEndingDirectorySeparator(System.IO.Path.GetFullPath(left));
        string normalizedRight = System.IO.Path.TrimEndingDirectorySeparator(System.IO.Path.GetFullPath(right));
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        return string.Equals(normalizedLeft, normalizedRight, comparison);
    }

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
}
