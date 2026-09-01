using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public enum GuidedWorkflowStep {
    Files,
    Options,
    Review,
    Run,
    Result
}

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
    [NotifyPropertyChangedFor(nameof(IsResultSuccessful))]
    [NotifyPropertyChangedFor(nameof(IsResultFailed))]
    [NotifyPropertyChangedFor(nameof(IsResultCancelled))]
    private OfficeWorkflowStatus? _resultStatus;

    [ObservableProperty]
    private bool _hasHealthReport;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsFilesStep))]
    [NotifyPropertyChangedFor(nameof(IsOptionsStep))]
    [NotifyPropertyChangedFor(nameof(IsReviewStep))]
    [NotifyPropertyChangedFor(nameof(IsRunStep))]
    [NotifyPropertyChangedFor(nameof(IsResultStep))]
    [NotifyPropertyChangedFor(nameof(IsFilesReached))]
    [NotifyPropertyChangedFor(nameof(IsOptionsReached))]
    [NotifyPropertyChangedFor(nameof(IsReviewReached))]
    [NotifyPropertyChangedFor(nameof(IsRunReached))]
    [NotifyPropertyChangedFor(nameof(IsResultReached))]
    [NotifyPropertyChangedFor(nameof(CanGoBack))]
    [NotifyPropertyChangedFor(nameof(CanContinue))]
    private GuidedWorkflowStep _currentStep;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanRun))]
    private bool _isBusy;

    public bool NeedsComparison => SelectedOperation.Value == OfficeWorkflowOperation.Compare;
    public bool ShowsOptimizationProfile => SelectedOperation.Value == OfficeWorkflowOperation.Optimize;
    public bool CanRun => !IsBusy && !string.IsNullOrWhiteSpace(InputPath) && (!NeedsComparison || !string.IsNullOrWhiteSpace(ComparisonPath));
    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(OutputPath);
    public bool IsResultSuccessful => ResultStatus == OfficeWorkflowStatus.Completed;
    public bool IsResultFailed => ResultStatus == OfficeWorkflowStatus.Failed;
    public bool IsResultCancelled => ResultStatus == OfficeWorkflowStatus.Cancelled;
    public bool IsFilesStep => CurrentStep == GuidedWorkflowStep.Files;
    public bool IsOptionsStep => CurrentStep == GuidedWorkflowStep.Options;
    public bool IsReviewStep => CurrentStep == GuidedWorkflowStep.Review;
    public bool IsRunStep => CurrentStep == GuidedWorkflowStep.Run;
    public bool IsResultStep => CurrentStep == GuidedWorkflowStep.Result;
    public bool IsFilesReached => CurrentStep >= GuidedWorkflowStep.Files;
    public bool IsOptionsReached => CurrentStep >= GuidedWorkflowStep.Options;
    public bool IsReviewReached => CurrentStep >= GuidedWorkflowStep.Review;
    public bool IsRunReached => CurrentStep >= GuidedWorkflowStep.Run;
    public bool IsResultReached => CurrentStep >= GuidedWorkflowStep.Result;
    public bool CanGoBack => !IsBusy && CurrentStep is GuidedWorkflowStep.Options or GuidedWorkflowStep.Review;
    public bool CanContinue => !IsBusy && (IsFilesStep
        ? !string.IsNullOrWhiteSpace(InputPath) && (!NeedsComparison || !string.IsNullOrWhiteSpace(ComparisonPath))
        : IsOptionsStep);
    public string InputFileName => string.IsNullOrWhiteSpace(InputPath) ? "No PDF selected" : Path.GetFileName(InputPath);
    public string InputDirectory => string.IsNullOrWhiteSpace(InputPath) ? "Choose the source document" : Path.GetDirectoryName(InputPath) ?? string.Empty;
    public string WorkbenchTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Inspect => "Inspect PDF",
        OfficeWorkflowOperation.Compare => "Compare PDFs",
        OfficeWorkflowOperation.Optimize => "Compress PDF",
        OfficeWorkflowOperation.RepairPlan => "Plan PDF repair",
        OfficeWorkflowOperation.Repair => "Repair PDF",
        OfficeWorkflowOperation.Sanitize => "Sanitize PDF",
        _ => "PDF workbench"
    };
    public string OperationVerb => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Inspect => "inspect",
        OfficeWorkflowOperation.Compare => "compare",
        OfficeWorkflowOperation.Optimize => "compress",
        OfficeWorkflowOperation.RepairPlan => "assess",
        OfficeWorkflowOperation.Repair => "repair",
        OfficeWorkflowOperation.Sanitize => "sanitize",
        _ => "process"
    };
    public string FilesDescription => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? "Select the two PDFs that OfficeIMO should compare. Both source files remain unchanged."
        : $"Select the PDF that OfficeIMO should {OperationVerb}. The source file remains unchanged.";
    public string OptionsTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Optimize => "Compression options",
        OfficeWorkflowOperation.Compare => "Comparison options",
        OfficeWorkflowOperation.Repair => "Repair options",
        _ => "Operation options"
    };
    public string ReviewFileLabel => SelectedOperation.Value == OfficeWorkflowOperation.Compare ? "Primary file" : "File to " + OperationVerb;
    public string PlanTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Optimize => "Compression plan",
        OfficeWorkflowOperation.Compare => "Comparison plan",
        OfficeWorkflowOperation.Repair => "Repair plan",
        _ => "Operation plan"
    };
    public string RunTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Repair => "Repairing and validating",
        OfficeWorkflowOperation.Optimize => "Compressing and validating",
        OfficeWorkflowOperation.Compare => "Comparing documents",
        OfficeWorkflowOperation.Inspect => "Inspecting document",
        OfficeWorkflowOperation.RepairPlan => "Assessing repairability",
        OfficeWorkflowOperation.Sanitize => "Sanitizing and validating",
        _ => "Running operation"
    };
    public string RunActionLabel => "Run " + OperationVerb;
    public string ArtifactAssuranceText => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Repair => "Repair publishes a separate artifact only after OfficeIMO can reopen and validate it.",
        OfficeWorkflowOperation.Optimize => "Compression publishes a separate artifact and retains the original when it is already smaller.",
        OfficeWorkflowOperation.Compare => "Comparison creates a separate evidence report without changing either PDF.",
        OfficeWorkflowOperation.Sanitize => "Sanitization publishes a separate artifact only after OfficeIMO validates the saved result.",
        _ => "The source PDF remains unchanged while OfficeIMO builds the operation report."
    };
    public string PlanStepOne => SelectedOperation.Value == OfficeWorkflowOperation.Compare ? "Structure comparison" : "Structure analysis";
    public string PlanStepOneDetail => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Compare => "Compare document structure and managed page renderings",
        OfficeWorkflowOperation.Optimize => "Inventory objects and select deterministic lossless rewrites",
        OfficeWorkflowOperation.Repair => "Deep scan and recover explicit structural defects",
        _ => "Inventory the document structure and supported capabilities"
    };
    public string PlanStepTwo => SelectedOperation.Value == OfficeWorkflowOperation.Compare ? "Evidence gallery" : "Content streams";
    public string PlanStepTwoDetail => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Compare => "Produce an HTML gallery for review",
        OfficeWorkflowOperation.Optimize => "Preserve content while reducing supported storage overhead",
        OfficeWorkflowOperation.Repair => "Preserve or safely rebuild recoverable streams",
        _ => "Preserve supported content and report unsupported paths"
    };
    public string PlanStepThree => "Output validation";
    public string PlanStepThreeDetail => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? "Record structural and visual differences"
        : "Strict reopen and preservation evidence";
    public string OutputPreviewPath {
        get {
            if (string.IsNullOrWhiteSpace(InputPath)) return "Choose a PDF to calculate the output path.";
            string input = Path.GetFullPath(InputPath);
            string directory = string.IsNullOrWhiteSpace(OutputFolder) ? Path.GetDirectoryName(input)! : Path.GetFullPath(OutputFolder);
            string stem = Path.Combine(directory, Path.GetFileNameWithoutExtension(input));
            return SelectedOperation.Value switch {
                OfficeWorkflowOperation.Compare => stem + ".comparison.html",
                OfficeWorkflowOperation.Optimize => stem + ".optimized.pdf",
                OfficeWorkflowOperation.Repair => stem + ".repaired.pdf",
                OfficeWorkflowOperation.Sanitize => stem + ".sanitized.pdf",
                _ => "No new PDF is written; results remain in the operation report."
            };
        }
    }

    partial void OnSelectedOperationChanged(HealthOperationChoice value) {
        OnPropertyChanged(nameof(NeedsComparison));
        OnPropertyChanged(nameof(ShowsOptimizationProfile));
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(WorkbenchTitle));
        OnPropertyChanged(nameof(OperationVerb));
        OnPropertyChanged(nameof(FilesDescription));
        OnPropertyChanged(nameof(OptionsTitle));
        OnPropertyChanged(nameof(ReviewFileLabel));
        OnPropertyChanged(nameof(PlanTitle));
        OnPropertyChanged(nameof(RunTitle));
        OnPropertyChanged(nameof(RunActionLabel));
        OnPropertyChanged(nameof(ArtifactAssuranceText));
        OnPropertyChanged(nameof(PlanStepOne));
        OnPropertyChanged(nameof(PlanStepOneDetail));
        OnPropertyChanged(nameof(PlanStepTwo));
        OnPropertyChanged(nameof(PlanStepTwoDetail));
        OnPropertyChanged(nameof(PlanStepThree));
        OnPropertyChanged(nameof(PlanStepThreeDetail));
        OnPropertyChanged(nameof(OutputPreviewPath));
        RunCommand.NotifyCanExecuteChanged();
    }

    partial void OnInputPathChanged(string value) {
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(CanContinue));
        OnPropertyChanged(nameof(InputFileName));
        OnPropertyChanged(nameof(InputDirectory));
        OnPropertyChanged(nameof(OutputPreviewPath));
        RunCommand.NotifyCanExecuteChanged();
        ContinueCommand.NotifyCanExecuteChanged();
    }

    partial void OnOutputFolderChanged(string value) => OnPropertyChanged(nameof(OutputPreviewPath));

    partial void OnComparisonPathChanged(string value) {
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(CanContinue));
        RunCommand.NotifyCanExecuteChanged();
        ContinueCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanCancel));
        OnPropertyChanged(nameof(CanGoBack));
        OnPropertyChanged(nameof(CanContinue));
        RunCommand.NotifyCanExecuteChanged();
        BackCommand.NotifyCanExecuteChanged();
        ContinueCommand.NotifyCanExecuteChanged();
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
        CurrentStep = GuidedWorkflowStep.Run;
        ProgressFraction = 0D;
        ResultStatus = null;
        HasHealthReport = false;
        OutputPath = null;
        Summary = "No report was produced.";
        BeforeSummary = "—";
        AfterSummary = "—";
        Diagnostics.Clear();
        Metrics.Clear();

        try {
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.Fraction;
                Status = update.Message;
            });
            OfficeWorkflowResult result = await _runner.RunAsync(CreateRequest(), progress, operationCancellation.Token).ConfigureAwait(true);
            ResultStatus = result.Status;
            Summary = result.Summary;
            OutputPath = result.OutputPath;
            Status = result.Status.ToString();
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) {
                Diagnostics.Add($"{diagnostic.Severity}: {diagnostic.Message}");
            }
            if (result.HealthReport is not null) {
                HasHealthReport = true;
                BeforeSummary = FormatSnapshot(result.HealthReport.Before);
                AfterSummary = result.HealthReport.After is null ? "No artifact was written." : FormatSnapshot(result.HealthReport.After);
                foreach ((string key, string value) in result.HealthReport.Metrics) Metrics.Add($"{FormatKey(key)}: {value}");
            }
        } finally {
            IsBusy = false;
            CurrentStep = GuidedWorkflowStep.Result;
            if (ReferenceEquals(_cancellation, operationCancellation)) _cancellation = null;
            operationCancellation.Dispose();
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    [RelayCommand(CanExecute = nameof(CanContinue))]
    private void Continue() {
        CurrentStep = CurrentStep switch {
            GuidedWorkflowStep.Files => GuidedWorkflowStep.Options,
            GuidedWorkflowStep.Options => GuidedWorkflowStep.Review,
            _ => CurrentStep
        };
        NotifyStepCommands();
    }

    [RelayCommand(CanExecute = nameof(CanGoBack))]
    private void Back() {
        CurrentStep = CurrentStep switch {
            GuidedWorkflowStep.Options => GuidedWorkflowStep.Files,
            GuidedWorkflowStep.Review => GuidedWorkflowStep.Options,
            _ => CurrentStep
        };
        NotifyStepCommands();
    }

    internal void PrepareRepairWorkflow() {
        PrepareWorkflow(OfficeWorkflowOperation.Repair);
    }

    internal void PrepareWorkflow(OfficeWorkflowOperation operation) {
        if (IsBusy) return;
        SelectedOperation = Operations.Single(choice => choice.Value == operation);
        CurrentStep = GuidedWorkflowStep.Files;
        Status = string.IsNullOrWhiteSpace(InputPath)
            ? $"Choose a PDF to {OperationVerb}."
            : $"Review the selected PDF and {OperationVerb} options.";
        NotifyStepCommands();
    }

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

    private void NotifyStepCommands() {
        OnPropertyChanged(nameof(IsFilesStep));
        OnPropertyChanged(nameof(IsOptionsStep));
        OnPropertyChanged(nameof(IsReviewStep));
        OnPropertyChanged(nameof(IsRunStep));
        OnPropertyChanged(nameof(IsResultStep));
        OnPropertyChanged(nameof(IsFilesReached));
        OnPropertyChanged(nameof(IsOptionsReached));
        OnPropertyChanged(nameof(IsReviewReached));
        OnPropertyChanged(nameof(IsRunReached));
        OnPropertyChanged(nameof(IsResultReached));
        OnPropertyChanged(nameof(CanGoBack));
        OnPropertyChanged(nameof(CanContinue));
        BackCommand.NotifyCanExecuteChanged();
        ContinueCommand.NotifyCanExecuteChanged();
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
