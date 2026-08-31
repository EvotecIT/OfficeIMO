using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class DocumentHealthViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeWorkflowRunner _runner;
    private CancellationTokenSource? _cancellation;

    public DocumentHealthViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner = null) {
        _pickPdf = pickPdf;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        SelectedOperation = Operations[0];
        SelectedProfile = Profiles[0];
    }

    public IReadOnlyList<HealthOperationChoice> Operations { get; } = [
        new(OfficeWorkflowOperation.Inspect, "Inspect", "Read structure, security, signatures, tags, active content, and repair diagnostics.", false),
        new(OfficeWorkflowOperation.Compare, "Compare", "Compare structure and managed page renderings, with an HTML review gallery.", true),
        new(OfficeWorkflowOperation.Optimize, "Optimize", "Apply a deterministic lossless profile and retain the original when it is smaller.", true),
        new(OfficeWorkflowOperation.RepairPlan, "Plan repair", "Assess recovered defects and blockers without writing a file.", false),
        new(OfficeWorkflowOperation.Repair, "Create repair artifact", "Persist explicit recoveries, reopen strictly, and prove preservation.", true),
        new(OfficeWorkflowOperation.Sanitize, "Sanitize", "Remove forbidden actions and payloads, then inventory the saved result.", true)
    ];

    public IReadOnlyList<WorkflowProfileChoice> Profiles { get; } = [
        new(OfficeWorkflowOutputProfile.Faithful, "Balanced", "Conservative deterministic lossless optimization."),
        new(OfficeWorkflowOutputProfile.Lightweight, "Maximum compression", "Use object and cross-reference streams where supported."),
        new(OfficeWorkflowOutputProfile.PrintReady, "Archival", "Classic cross references without linearization."),
        new(OfficeWorkflowOutputProfile.TextOnly, "Web", "Fast Web View layout with broadly compatible cross references.")
    ];

    public ObservableCollection<string> Diagnostics { get; } = new();
    public ObservableCollection<string> Metrics { get; } = new();

    [ObservableProperty]
    private HealthOperationChoice _selectedOperation = null!;

    [ObservableProperty]
    private WorkflowProfileChoice _selectedProfile = null!;

    [ObservableProperty]
    private string _inputPath = string.Empty;

    [ObservableProperty]
    private string _comparisonPath = string.Empty;

    [ObservableProperty]
    private string _outputFolder = string.Empty;

    [ObservableProperty]
    private string _pdfPassword = string.Empty;

    [ObservableProperty]
    private string _status = "Choose a PDF and an operation.";

    [ObservableProperty]
    private string _summary = "No report yet.";

    [ObservableProperty]
    private string _beforeSummary = "—";

    [ObservableProperty]
    private string _afterSummary = "—";

    [ObservableProperty]
    private string? _outputPath;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanRun))]
    private bool _isBusy;

    public bool NeedsComparison => SelectedOperation.Value == OfficeWorkflowOperation.Compare;
    public bool ShowsOptimizationProfile => SelectedOperation.Value == OfficeWorkflowOperation.Optimize;
    public bool CanRun => !IsBusy && !string.IsNullOrWhiteSpace(InputPath) && (!NeedsComparison || !string.IsNullOrWhiteSpace(ComparisonPath));
    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(OutputPath);

    partial void OnSelectedOperationChanged(HealthOperationChoice value) {
        OnPropertyChanged(nameof(NeedsComparison));
        OnPropertyChanged(nameof(ShowsOptimizationProfile));
        OnPropertyChanged(nameof(CanRun));
        RunCommand.NotifyCanExecuteChanged();
    }

    partial void OnInputPathChanged(string value) {
        OnPropertyChanged(nameof(CanRun));
        RunCommand.NotifyCanExecuteChanged();
    }

    partial void OnComparisonPathChanged(string value) {
        OnPropertyChanged(nameof(CanRun));
        RunCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanCancel));
        RunCommand.NotifyCanExecuteChanged();
    }

    partial void OnOutputPathChanged(string? value) => OnPropertyChanged(nameof(HasOutput));

    [RelayCommand]
    private async Task ChooseInputAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) InputPath = path;
    }

    [RelayCommand]
    private async Task ChooseComparisonAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) ComparisonPath = path;
    }

    [RelayCommand]
    private async Task ChooseOutputFolderAsync(CancellationToken cancellationToken) {
        string? path = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) OutputFolder = path;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync() {
        _cancellation?.Dispose();
        var operationCancellation = new CancellationTokenSource();
        _cancellation = operationCancellation;
        IsBusy = true;
        ProgressFraction = 0D;
        OutputPath = null;
        Diagnostics.Clear();
        Metrics.Clear();

        try {
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.Fraction;
                Status = update.Message;
            });
            OfficeWorkflowResult result = await _runner.RunAsync(CreateRequest(), progress, operationCancellation.Token).ConfigureAwait(true);
            Summary = result.Summary;
            OutputPath = result.OutputPath;
            Status = result.Status.ToString();
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) {
                Diagnostics.Add($"{diagnostic.Severity}: {diagnostic.Message}");
            }
            if (result.HealthReport is not null) {
                BeforeSummary = FormatSnapshot(result.HealthReport.Before);
                AfterSummary = result.HealthReport.After is null ? "No artifact was written." : FormatSnapshot(result.HealthReport.After);
                foreach ((string key, string value) in result.HealthReport.Metrics) Metrics.Add($"{FormatKey(key)}: {value}");
            }
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

    private OfficeWorkflowRequest CreateRequest() {
        string input = Path.GetFullPath(InputPath);
        string directory = string.IsNullOrWhiteSpace(OutputFolder) ? Path.GetDirectoryName(input)! : Path.GetFullPath(OutputFolder);
        string? output = SelectedOperation.Value switch {
            OfficeWorkflowOperation.Compare => Path.Combine(directory, Path.GetFileNameWithoutExtension(input) + ".comparison.html"),
            OfficeWorkflowOperation.Optimize => Path.Combine(directory, Path.GetFileNameWithoutExtension(input) + ".optimized.pdf"),
            OfficeWorkflowOperation.Repair => Path.Combine(directory, Path.GetFileNameWithoutExtension(input) + ".repaired.pdf"),
            OfficeWorkflowOperation.Sanitize => Path.Combine(directory, Path.GetFileNameWithoutExtension(input) + ".sanitized.pdf"),
            _ => null
        };
        return new OfficeWorkflowRequest {
            Operation = SelectedOperation.Value,
            InputPath = input,
            ComparisonPath = NeedsComparison ? Path.GetFullPath(ComparisonPath) : null,
            OutputPath = output,
            OutputProfile = SelectedProfile.Value,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Rename,
            PdfPassword = string.IsNullOrEmpty(PdfPassword) ? null : PdfPassword
        };
    }

    private static string FormatSnapshot(PdfHealthSnapshot snapshot) {
        var flags = new List<string>();
        if (snapshot.HasEncryption) flags.Add("encrypted");
        if (snapshot.HasSignatures) flags.Add("signed");
        if (snapshot.HasTaggedContent) flags.Add("tagged");
        if (snapshot.HasActiveContent) flags.Add("active content");
        if (snapshot.HasEmbeddedFiles) flags.Add("attachments");
        string features = flags.Count == 0 ? "no security or active-content markers" : string.Join(", ", flags);
        return $"{snapshot.PageCount:N0} {(snapshot.PageCount == 1 ? "page" : "pages")} · {FormatBytes(snapshot.SizeBytes)} · PDF {snapshot.Version ?? "?"}\n" +
               $"Read: {(snapshot.CanRead ? "yes" : "blocked")} · General rewrite: {(snapshot.CanRewrite ? "yes" : "blocked")} · {features}\n" +
               $"Recovered defects: {snapshot.RepairCount:N0} · Detected only: {snapshot.DetectionOnlyCount:N0}";
    }

    private static string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024D && unit < units.Length - 1) { value /= 1024D; unit++; }
        return value.ToString("0.#", CultureInfo.InvariantCulture) + " " + units[unit];
    }

    private static string FormatKey(string key) => string.Concat(key.Select((character, index) =>
        index > 0 && char.IsUpper(character) ? " " + char.ToLowerInvariant(character) : character.ToString()));
}
