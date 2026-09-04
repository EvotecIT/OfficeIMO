using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Localization;
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
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _cancellation;

    public DocumentHealthViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner = null) : this(pickPdf, pickOutputFolder, runner, null) { }

    internal DocumentHealthViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeWorkflowRunner? runner,
        IStudioLocalizer? localizer = null) {
        _pickPdf = pickPdf;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        _localizer = localizer ?? StudioLocalization.Current;
        Operations = [
            Operation(OfficeWorkflowOperation.Inspect, "Inspect", "Read structure, security, signatures, tags, active content, and repair diagnostics.", false),
            Operation(OfficeWorkflowOperation.Compare, "Compare", "Compare structure and managed page renderings, with an HTML review gallery.", true),
            Operation(OfficeWorkflowOperation.Optimize, "Optimize", "Apply a deterministic lossless profile and retain the original when it is smaller.", true),
            Operation(OfficeWorkflowOperation.RepairPlan, "Plan repair", "Assess recovered defects and blockers without writing a file.", false),
            Operation(OfficeWorkflowOperation.Repair, "Create repair artifact", "Persist explicit recoveries, reopen strictly, and prove preservation.", true),
            Operation(OfficeWorkflowOperation.Sanitize, "Sanitize", "Remove forbidden actions and payloads, then inventory the saved result.", true)
        ];
        Profiles = [
            Profile(OfficeWorkflowOutputProfile.Faithful, "Balanced", "Conservative deterministic lossless optimization."),
            Profile(OfficeWorkflowOutputProfile.Lightweight, "Maximum compression", "Use object and cross-reference streams where supported."),
            Profile(OfficeWorkflowOutputProfile.PrintReady, "Archival", "Classic cross references without linearization."),
            Profile(OfficeWorkflowOutputProfile.TextOnly, "Web", "Fast Web View layout with broadly compatible cross references.")
        ];
        SelectedOperation = Operations[0];
        SelectedProfile = Profiles[0];
        Status = T("Status.Ready", "Choose a PDF and an operation.");
        Summary = T("Summary.Empty", "No report yet.");
    }

    public IReadOnlyList<HealthOperationChoice> Operations { get; }

    public IReadOnlyList<WorkflowProfileChoice> Profiles { get; }

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
    private string _status = string.Empty;

    [ObservableProperty]
    private string _summary = string.Empty;

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
    public string InputFileName => string.IsNullOrWhiteSpace(InputPath) ? T("Input.None", "No PDF selected") : Path.GetFileName(InputPath);
    public string InputDirectory => string.IsNullOrWhiteSpace(InputPath) ? T("Input.Choose", "Choose the source document") : Path.GetDirectoryName(InputPath) ?? string.Empty;
    public string WorkbenchTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Inspect => T("Title.Inspect", "Inspect PDF"),
        OfficeWorkflowOperation.Compare => T("Title.Compare", "Compare PDFs"),
        OfficeWorkflowOperation.Optimize => T("Title.Optimize", "Compress PDF"),
        OfficeWorkflowOperation.RepairPlan => T("Title.RepairPlan", "Plan PDF repair"),
        OfficeWorkflowOperation.Repair => T("Title.Repair", "Repair PDF"),
        OfficeWorkflowOperation.Sanitize => T("Title.Sanitize", "Sanitize PDF"),
        _ => T("Title.Default", "PDF workbench")
    };
    public string OperationVerb => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Inspect => T("Verb.Inspect", "inspect"),
        OfficeWorkflowOperation.Compare => T("Verb.Compare", "compare"),
        OfficeWorkflowOperation.Optimize => T("Verb.Optimize", "compress"),
        OfficeWorkflowOperation.RepairPlan => T("Verb.RepairPlan", "assess"),
        OfficeWorkflowOperation.Repair => T("Verb.Repair", "repair"),
        OfficeWorkflowOperation.Sanitize => T("Verb.Sanitize", "sanitize"),
        _ => T("Verb.Default", "process")
    };
    public string FilesDescription => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? T("Files.Compare", "Select the two PDFs that OfficeIMO should compare. Both source files remain unchanged.")
        : _localizer.FormatOrDefault($"Health.Files.{SelectedOperation.Value}", "Select the PDF that OfficeIMO should {0}. The source file remains unchanged.", OperationVerb);
    public string OptionsTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Optimize => T("Options.Optimize", "Compression options"),
        OfficeWorkflowOperation.Compare => T("Options.Compare", "Comparison options"),
        OfficeWorkflowOperation.Repair => T("Options.Repair", "Repair options"),
        _ => T("Options.Default", "Operation options")
    };
    public string ReviewFileLabel => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? T("Review.PrimaryFile", "Primary file")
        : _localizer.FormatOrDefault($"Health.Review.File.{SelectedOperation.Value}", "File to {0}", OperationVerb);
    public string PlanTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Optimize => T("Plan.Optimize", "Compression plan"),
        OfficeWorkflowOperation.Compare => T("Plan.Compare", "Comparison plan"),
        OfficeWorkflowOperation.Repair => T("Plan.Repair", "Repair plan"),
        _ => T("Plan.Default", "Operation plan")
    };
    public string RunTitle => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Repair => T("Run.Repair", "Repairing and validating"),
        OfficeWorkflowOperation.Optimize => T("Run.Optimize", "Compressing and validating"),
        OfficeWorkflowOperation.Compare => T("Run.Compare", "Comparing documents"),
        OfficeWorkflowOperation.Inspect => T("Run.Inspect", "Inspecting document"),
        OfficeWorkflowOperation.RepairPlan => T("Run.RepairPlan", "Assessing repairability"),
        OfficeWorkflowOperation.Sanitize => T("Run.Sanitize", "Sanitizing and validating"),
        _ => T("Run.Default", "Running operation")
    };
    public string RunActionLabel => _localizer.FormatOrDefault($"Health.RunAction.{SelectedOperation.Value}", "Run {0}", OperationVerb);
    public string ArtifactAssuranceText => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Repair => T("Assurance.Repair", "Repair publishes a separate artifact only after OfficeIMO can reopen and validate it."),
        OfficeWorkflowOperation.Optimize => T("Assurance.Optimize", "Compression publishes a separate artifact and retains the original when it is already smaller."),
        OfficeWorkflowOperation.Compare => T("Assurance.Compare", "Comparison creates a separate evidence report without changing either PDF."),
        OfficeWorkflowOperation.Sanitize => T("Assurance.Sanitize", "Sanitization publishes a separate artifact only after OfficeIMO validates the saved result."),
        _ => T("Assurance.Default", "The source PDF remains unchanged while OfficeIMO builds the operation report.")
    };
    public string PlanStepOne => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? T("Plan.Step1.Compare", "Structure comparison")
        : T("Plan.Step1.Default", "Structure analysis");
    public string PlanStepOneDetail => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Compare => T("Plan.Step1Detail.Compare", "Compare document structure and managed page renderings"),
        OfficeWorkflowOperation.Optimize => T("Plan.Step1Detail.Optimize", "Inventory objects and select deterministic lossless rewrites"),
        OfficeWorkflowOperation.Repair => T("Plan.Step1Detail.Repair", "Deep scan and recover explicit structural defects"),
        _ => T("Plan.Step1Detail.Default", "Inventory the document structure and supported capabilities")
    };
    public string PlanStepTwo => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? T("Plan.Step2.Compare", "Evidence gallery")
        : T("Plan.Step2.Default", "Content streams");
    public string PlanStepTwoDetail => SelectedOperation.Value switch {
        OfficeWorkflowOperation.Compare => T("Plan.Step2Detail.Compare", "Produce an HTML gallery for review"),
        OfficeWorkflowOperation.Optimize => T("Plan.Step2Detail.Optimize", "Preserve content while reducing supported storage overhead"),
        OfficeWorkflowOperation.Repair => T("Plan.Step2Detail.Repair", "Preserve or safely rebuild recoverable streams"),
        _ => T("Plan.Step2Detail.Default", "Preserve supported content and report unsupported paths")
    };
    public string PlanStepThree => T("Plan.Step3", "Output validation");
    public string PlanStepThreeDetail => SelectedOperation.Value == OfficeWorkflowOperation.Compare
        ? T("Plan.Step3Detail.Compare", "Record structural and visual differences")
        : T("Plan.Step3Detail.Default", "Strict reopen and preservation evidence");
    public string OutputPreviewPath {
        get {
            if (string.IsNullOrWhiteSpace(InputPath)) return T("Output.ChooseInput", "Choose a PDF to calculate the output path.");
            string input = Path.GetFullPath(InputPath);
            string directory = string.IsNullOrWhiteSpace(OutputFolder) ? Path.GetDirectoryName(input)! : Path.GetFullPath(OutputFolder);
            string stem = Path.Combine(directory, Path.GetFileNameWithoutExtension(input));
            return SelectedOperation.Value switch {
                OfficeWorkflowOperation.Compare => stem + ".comparison.html",
                OfficeWorkflowOperation.Optimize => stem + ".optimized.pdf",
                OfficeWorkflowOperation.Repair => stem + ".repaired.pdf",
                OfficeWorkflowOperation.Sanitize => stem + ".sanitized.pdf",
                _ => T("Output.NoArtifact", "No new PDF is written; results remain in the operation report.")
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
        Summary = T("Summary.NotProduced", "No report was produced.");
        BeforeSummary = "—";
        AfterSummary = "—";
        Diagnostics.Clear();
        Metrics.Clear();

        try {
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.Fraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
            });
            OfficeWorkflowResult result = await _runner.RunAsync(CreateRequest(), progress, operationCancellation.Token).ConfigureAwait(true);
            ResultStatus = result.Status;
            Summary = result.Summary;
            OutputPath = result.OutputPath;
            Status = _localizer.GetOrDefault($"Workflow.Status.{result.Status}", result.Status.ToString());
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) {
                Diagnostics.Add(_localizer.FormatOrDefault(
                    $"Workflow.Diagnostic.{diagnostic.Severity}",
                    "{0}: {1}",
                    _localizer.GetOrDefault($"Workflow.Severity.{diagnostic.Severity}", diagnostic.Severity.ToString()),
                    diagnostic.Message));
            }
            if (result.HealthReport is not null) {
                HasHealthReport = true;
                BeforeSummary = FormatSnapshot(result.HealthReport.Before);
                AfterSummary = result.HealthReport.After is null ? T("Output.NotWritten", "No artifact was written.") : FormatSnapshot(result.HealthReport.After);
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
            ? _localizer.FormatOrDefault($"Health.Prepare.Empty.{operation}", "Choose a PDF to {0}.", OperationVerb)
            : _localizer.FormatOrDefault($"Health.Prepare.Ready.{operation}", "Review the selected PDF and {0} options.", OperationVerb);
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

    private string FormatSnapshot(PdfHealthSnapshot snapshot) {
        var flags = new List<string>();
        if (snapshot.HasEncryption) flags.Add(T("Snapshot.Encrypted", "encrypted"));
        if (snapshot.HasSignatures) flags.Add(T("Snapshot.Signed", "signed"));
        if (snapshot.HasTaggedContent) flags.Add(T("Snapshot.Tagged", "tagged"));
        if (snapshot.HasActiveContent) flags.Add(T("Snapshot.ActiveContent", "active content"));
        if (snapshot.HasEmbeddedFiles) flags.Add(T("Snapshot.Attachments", "attachments"));
        string features = flags.Count == 0 ? T("Snapshot.NoMarkers", "no security or active-content markers") : string.Join(", ", flags);
        return _localizer.FormatOrDefault(
            "Health.Snapshot",
            "{0:N0} {1} · {2} · PDF {3}\nRead: {4} · General rewrite: {5} · {6}\nRecovered defects: {7:N0} · Detected only: {8:N0}",
            snapshot.PageCount,
            snapshot.PageCount == 1 ? T("Snapshot.Page", "page") : T("Snapshot.Pages", "pages"),
            FormatBytes(snapshot.SizeBytes),
            snapshot.Version ?? "?",
            snapshot.CanRead ? T("Snapshot.Yes", "yes") : T("Snapshot.Blocked", "blocked"),
            snapshot.CanRewrite ? T("Snapshot.Yes", "yes") : T("Snapshot.Blocked", "blocked"),
            features,
            snapshot.RepairCount,
            snapshot.DetectionOnlyCount);
    }

    private string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024D && unit < units.Length - 1) { value /= 1024D; unit++; }
        return value.ToString("0.#", _localizer.Culture) + " " + units[unit];
    }

    private static string FormatKey(string key) => string.Concat(key.Select((character, index) =>
        index > 0 && char.IsUpper(character) ? " " + char.ToLowerInvariant(character) : character.ToString()));

    private HealthOperationChoice Operation(OfficeWorkflowOperation value, string label, string description, bool producesArtifact) =>
        new(value, T($"Operation.{value}.Label", label), T($"Operation.{value}.Description", description), producesArtifact);

    private WorkflowProfileChoice Profile(OfficeWorkflowOutputProfile value, string label, string description) =>
        new(value, T($"Profile.{value}.Label", label), T($"Profile.{value}.Description", description));

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Health." + suffix, fallback);
}
