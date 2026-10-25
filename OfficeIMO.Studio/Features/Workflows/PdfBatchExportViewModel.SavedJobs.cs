using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class PdfBatchExportViewModel {
    private readonly Func<IStudioFileDialogs> _fileDialogs;
    private OfficeConversionJob? _loadedJob;
    private string _jobBaseDirectory = Path.GetFullPath(Environment.CurrentDirectory);
    private bool _applyingJob, _savedMatchesFields, _hasRun;
    private Func<string, Task>? _copyCommandText;

    /// <summary>Host-owned clipboard operation; saved jobs contain no clipboard or platform logic.</summary>
    internal Func<string, Task>? CopyCommandText {
        get => _copyCommandText;
        set { _copyCommandText = value; RefreshSavedJobCommands(); }
    }

    [ObservableProperty] private string _savedJobPath = string.Empty;
    [ObservableProperty] private string _planSummary = string.Empty;
    [ObservableProperty] private string _jobSettingsSummary = string.Empty;
    [ObservableProperty] private string _planCountSummary = string.Empty;
    [ObservableProperty] private string _planFindings = string.Empty;
    [ObservableProperty] private IReadOnlyList<SavedFolderJobPlanItem> _planItems = [];
    public bool HasPlanSummary => !string.IsNullOrEmpty(PlanSummary);
    public bool HasJobSettingsSummary => !string.IsNullOrEmpty(JobSettingsSummary);
    public bool HasSavedJob => !string.IsNullOrEmpty(SavedJobPath);
    public bool HasPlanFindings => !string.IsNullOrEmpty(PlanFindings);
    public string SavedJobName => HasSavedJob ? Path.GetFileName(SavedJobPath) : T("Unsaved", "Unsaved folder job");
    public string InputSummary => string.IsNullOrWhiteSpace(InputDirectory) ? T("ChooseInput", "Choose the source folder on the left.") : InputDirectory;
    public string OutputSummary => string.IsNullOrWhiteSpace(OutputDirectory) ? T("ChooseOutput", "Choose the PDF output folder on the left.") : OutputDirectory;
    public string CheckpointSummary => string.IsNullOrWhiteSpace(CheckpointDirectory) ? T("NoCheckpoints", "No checkpoints; each run follows the saved conflict policy.") : CheckpointDirectory;
    public bool CanCopySavedJob => CanEdit && _savedMatchesFields && HasSavedJob && _copyCommandText is not null;
    public string RunLabel => T(_hasRun ? "RunAgain" : "Run", _hasRun ? "Run again" : "Run folder export");
    public string SavedJobHint => _savedMatchesFields
        ? T("SavedReady", "Saved settings are ready to use in Studio, the command line or PowerShell. Loading never starts work.")
        : T("SaveBeforeCopy", "Save these settings to copy a command. Running again verifies checkpointed outputs before reusing them.");

    partial void OnPlanSummaryChanged(string value) => OnPropertyChanged(nameof(HasPlanSummary));
    partial void OnJobSettingsSummaryChanged(string value) => OnPropertyChanged(nameof(HasJobSettingsSummary));
    partial void OnPlanFindingsChanged(string value) => OnPropertyChanged(nameof(HasPlanFindings));
    partial void OnSavedJobPathChanged(string value) { OnPropertyChanged(nameof(HasSavedJob)); OnPropertyChanged(nameof(SavedJobName)); }
    partial void OnInputDirectoryChanged(string value) { OnPropertyChanged(nameof(InputSummary)); InvalidateJobReview(); }
    partial void OnOutputDirectoryChanged(string value) { OnPropertyChanged(nameof(OutputSummary)); InvalidateJobReview(); }
    partial void OnCheckpointDirectoryChanged(string value) { OnPropertyChanged(nameof(CheckpointSummary)); InvalidateJobReview(); }
    partial void OnAllowLegacyImportLossChanged(bool value) => InvalidateJobReview();
    partial void OnRetryFailedChanged(bool value) => InvalidateJobReview();
    partial void OnTextEncodingChanged(string value) => InvalidateJobReview();
    partial void OnTabSizeChanged(decimal value) => InvalidateJobReview();

    private void InvalidateJobReview() {
        if (_applyingJob) return;
        _savedMatchesFields = false;
        ClearPlanReview();
        RefreshSavedJobCommands();
    }

    private void ClearPlanReview() { PlanSummary = PlanCountSummary = PlanFindings = string.Empty; PlanItems = []; }

    private void RefreshSavedJobCommands() {
        OnPropertyChanged(nameof(CanCopySavedJob));
        OnPropertyChanged(nameof(SavedJobHint));
        SaveJobCommand.NotifyCanExecuteChanged(); LoadJobCommand.NotifyCanExecuteChanged();
        PlanJobCommand.NotifyCanExecuteChanged(); CopyCliCommand.NotifyCanExecuteChanged(); CopyPowerShellCommand.NotifyCanExecuteChanged();
    }

    /// <summary>Changes only the fields this surface edits; every other loaded portable setting remains intact.</summary>
    private OfficeConversionJob CreateCurrentJob() {
        OfficeConversionBatchRequest request = _loadedJob?.CreateRequest(_jobBaseDirectory)
            ?? new OfficeConversionBatchRequest { OutputDirectory = OutputDirectory };
        OfficeWorkflowConversionOptions options = request.ConversionOptions.Clone();
        options.LegacyDocLossPolicy = AllowLegacyImportLoss ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block;
        var text = options.PlainText?.Clone() ?? new OfficeIMO.Pdf.PdfPlainTextOptions();
        text.EncodingName = string.IsNullOrWhiteSpace(TextEncoding) ? null : TextEncoding;
        text.TabSize = (int)TabSize;
        options.PlainText = text;
        request = request with {
            InputDirectory = InputDirectory, OutputDirectory = OutputDirectory,
            CheckpointDirectory = string.IsNullOrWhiteSpace(CheckpointDirectory) ? null : CheckpointDirectory,
            RetryFailed = RetryFailed, ConversionOptions = options
        };
        var definition = OfficeConversionJob.FromRequest(request, _jobBaseDirectory);
        definition.Name = _loadedJob?.Name;
        return definition;
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task SaveJobAsync(CancellationToken token) => WithJobOperationAsync(async cancellation => {
        OfficeConversionBatchRequest request = CreateCurrentJob().CreateRequest(_jobBaseDirectory);
        string? location = await _fileDialogs().PickSaveFileAsync(T("SaveTitle", "Save folder conversion job"),
            HasSavedJob ? Path.GetFileName(SavedJobPath) : "conversion.oijob.json",
            new StudioFileType(T("FileType", "OfficeIMO conversion job"), ["json"], "application/json"), cancellation);
        if (location is null) return;
        string path = RequireLocalJobPath(location);
        string baseDirectory = Path.GetDirectoryName(path)!;
        OfficeConversionJob definition = OfficeConversionJob.FromRequest(request, baseDirectory);
        definition.Name = _loadedJob?.Name;
        string serialized = OfficeConversionJobSerializer.Serialize(definition);
        OfficeConversionJobSerializer.ValidateStoragePath(definition, path);
        await AuthorizeJobDestinationAsync(path, request, cancellation);
        string? expected = null;
        if (File.Exists(path)) {
            // Replacing an existing job is useful; replacing an unrelated document is not a save-job operation.
            await Task.Run(() => OfficeConversionJobSerializer.ReadFile(path), cancellation);
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            expected = Convert.ToHexString(await System.Security.Cryptography.SHA256.HashDataAsync(stream, cancellation));
        }
        await StudioLocalPublication.WriteVerifiedAsync(path, Encoding.UTF8.GetBytes(serialized), expected,
            ct => AuthorizeJobDestinationAsync(path, request, ct), cancellation, _localizer);
        SavedJobPath = path;
        _loadedJob = definition; _jobBaseDirectory = baseDirectory;
        _applyingJob = true;
        try { InputDirectory = request.InputDirectory!; OutputDirectory = request.OutputDirectory; CheckpointDirectory = request.CheckpointDirectory ?? string.Empty; }
        finally { _applyingJob = false; }
        _savedMatchesFields = true;
        JobSettingsSummary = DescribeSettings(request);
        Status = T("Saved", "Folder job saved and verified. No documents were processed.");
    }, token);

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task LoadJobAsync(CancellationToken token) => WithJobOperationAsync(async cancellation => {
        string? location = await _fileDialogs().PickOpenFileAsync(T("LoadTitle", "Load folder conversion job"),
            new StudioFileType(T("FileType", "OfficeIMO conversion job"), ["json"], "application/json"), cancellation);
        if (location is null) return;
        string path = RequireLocalJobPath(location);
        OfficeConversionJob definition = await Task.Run(() => OfficeConversionJobSerializer.ReadFile(path), cancellation);
        if (definition.Secrets.SourcePasswordEnvironmentVariable is not null || definition.Secrets.PdfPasswordEnvironmentVariable is not null)
            throw new NotSupportedException(T("Secrets", "This workbench cannot request saved-job password references. Run that job through the CLI or PowerShell."));
        string baseDirectory = Path.GetDirectoryName(path)!;
        OfficeConversionBatchRequest request = definition.CreateRequest(baseDirectory);
        ValidateFolderJob(request);
        // Validation and projection finish before any visible fields or the retained definition change.
        string summary = DescribeSettings(request);
        _applyingJob = true;
        try {
            InputDirectory = request.InputDirectory!; OutputDirectory = request.OutputDirectory;
            CheckpointDirectory = request.CheckpointDirectory ?? string.Empty; RetryFailed = request.RetryFailed;
            AllowLegacyImportLoss = request.ConversionOptions.LegacyDocLossPolicy == OfficeConversionLossPolicy.Allow;
            TextEncoding = request.ConversionOptions.PlainText?.EncodingName ?? string.Empty;
            TabSize = request.ConversionOptions.PlainText?.TabSize ?? 8;
            _loadedJob = definition; _jobBaseDirectory = baseDirectory;
            SavedJobPath = path; JobSettingsSummary = summary; ClearPlanReview();
            _savedMatchesFields = true; _hasRun = false; OnPropertyChanged(nameof(RunLabel));
        } finally { _applyingJob = false; }
        Status = T("Loaded", "Folder job loaded. Review its settings and plan before running; no documents were processed.");
    }, token);

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private Task PlanJobAsync(CancellationToken token) => WithJobOperationAsync(async cancellation => {
        OfficeConversionJob definition = CreateCurrentJob();
        var plan = await OfficeConversionJobRunner.PlanAsync(definition, _jobBaseDirectory, _runner,
            cancellationToken: cancellation);
        PlanSummary = _localizer.FormatOrDefault("Conversion.SavedJob.PlanResult",
            "Found {0:N0} · Selected {1:N0} · Skipped {2:N0} · Failed {3:N0} · Output conflicts {4:N0}",
            plan.Discovered, plan.Selected, plan.Skipped, plan.Failed, plan.Collisions);
        PlanCountSummary = PlanSummary;
        PlanItems = plan.Items.Take(12).Select(item => new SavedFolderJobPlanItem(Path.GetFileName(item.InputPath), item.InputPath,
            item.OutputPath ?? "—", item.Message)).ToArray();
        PlanFindings = string.Join(Environment.NewLine, plan.Diagnostics);
        var findings = plan.Items.Take(12).Select(item => Path.GetFileName(item.InputPath) + " → " +
            (item.OutputPath ?? "—") + Environment.NewLine + item.Message);
        PlanSummary += Environment.NewLine + string.Join(Environment.NewLine, plan.Diagnostics.Concat(findings));
        if (plan.ItemsTruncated || plan.Items.Count > 12) {
            string sample = T("PlanSample", "Showing the first plan items; the counts cover the whole selection.");
            PlanSummary += Environment.NewLine + sample; PlanFindings += Environment.NewLine + sample;
        }
        JobSettingsSummary = DescribeSettings(definition.CreateRequest(_jobBaseDirectory));
        Status = plan.CanRun ? T("PlanReady", "Plan ready. No output or checkpoints were written. Review the destinations before running.")
            : T("PlanBlocked", "The plan has findings to resolve before running. No output or checkpoints were written.");
    }, token);

    [RelayCommand(CanExecute = nameof(CanCopySavedJob))]
    private Task CopyCliAsync(CancellationToken token) => CopySavedCommandAsync(false, token);
    [RelayCommand(CanExecute = nameof(CanCopySavedJob))]
    private Task CopyPowerShellAsync(CancellationToken token) => CopySavedCommandAsync(true, token);

    private Task CopySavedCommandAsync(bool powerShell, CancellationToken token) => WithJobOperationAsync(async cancellation => {
        // A command points at the saved file, so do not copy it after that file changed outside Studio.
        OfficeConversionJob current = await Task.Run(() => OfficeConversionJobSerializer.ReadFile(SavedJobPath), cancellation);
        if (_loadedJob is null || OfficeConversionJobSerializer.Serialize(current) != OfficeConversionJobSerializer.Serialize(_loadedJob))
            throw new IOException(T("FileChanged", "The saved job changed outside Studio. Load it again or save the reviewed settings before copying a command."));
        await _copyCommandText!(powerShell ? OfficeConversionJobCommands.GetPowerShellCommand(SavedJobPath)
            : OfficeConversionJobCommands.GetCliCommand(SavedJobPath));
        Status = T("Copied", "Command copied for the saved job. It runs only when you execute it.");
    }, token);

    private async Task WithJobOperationAsync(Func<CancellationToken, Task> operation, CancellationToken token) {
        if (!CanEdit) return;
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _cancellation = cancellation; IsBusy = true;
        try { await operation(cancellation.Token); }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { Status = T("Cancelled", "Job action cancelled. Existing documents and outputs were retained."); }
        catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) { Status = error.Message; }
        finally { _cancellation = null; IsBusy = false; RefreshSavedJobCommands(); }
    }

    private string RequireLocalJobPath(string location) => OfficeStorageIdentity.GetLocalPath(location)
        ?? throw new IOException(T("LocalFile", "Choose a local job file. Saved folder jobs need filesystem paths that Studio, the CLI and PowerShell can all use."));

    private void ValidateFolderJob(OfficeConversionBatchRequest request) {
        if (string.IsNullOrWhiteSpace(request.InputDirectory) || request.InputPaths is { Length: > 0 } ||
            !string.Equals(request.TargetExtension, ".pdf", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrWhiteSpace(request.ConversionRouteId))
            throw new NotSupportedException(T("FolderOnly", "This workbench loads folder-to-PDF jobs. Use the CLI or PowerShell for explicit-file lists, other targets or a fixed conversion route."));
        if (!string.IsNullOrEmpty(request.PdfPassword) || !string.IsNullOrEmpty(request.ConversionOptions.SourcePassword))
            throw new NotSupportedException(T("Secrets", "This workbench cannot request saved-job password references. Run that job through the CLI or PowerShell."));
    }

    private async Task AuthorizeJobDestinationAsync(string path, OfficeConversionBatchRequest request, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeConversionJobSerializer.ValidateStoragePath(OfficeConversionJob.FromRequest(request, Path.GetDirectoryName(path)!), path);
        if (OfficeStorageIdentity.AreEquivalent(path, request.InputDirectory!) ||
            request.InputPaths?.Any(input => OfficeStorageIdentity.AreEquivalent(path, input)) == true)
            throw new IOException(T("PreserveInput", "Choose a separate job filename to preserve its inputs."));
        if (_guard is not null && !await _guard.CanPublishAsync(path, false, token))
            throw new IOException(T("Protected", "The job destination is used by an open document or recovery operation."));
    }

    private string DescribeSettings(OfficeConversionBatchRequest request) {
        string summary = _localizer.FormatOrDefault("Conversion.SavedJob.Settings",
        "PDF · {0} profile · {1} conflicts · {2} concurrent · {3:N0} files maximum · {4:N0} MiB input / {5:N0} MiB output · {6}.",
        request.OutputProfile, request.ConflictPolicy, request.MaximumConcurrency, request.MaximumFiles,
        request.MaximumInputBytes / (1024 * 1024), request.MaximumOutputBytes / (1024 * 1024),
        request.Recursive ? T("Recursive", "includes subfolders") : T("TopFolder", "top folder only"));
        string filters = request.SourceExtensions is { Length: > 0 } ? string.Join(", ", request.SourceExtensions) : T("AllFormats", "all supported formats");
        summary += Environment.NewLine + _localizer.FormatOrDefault("Conversion.SavedJob.Filters", "Source selection: {0}. XML limit: {1:N0} characters per part.", filters, request.MaximumXmlCharactersInPart);
        var options = request.ConversionOptions;
        var extra = new List<string>();
        void Add(string name, string fallback, object? value) {
            if (value is not null) extra.Add(_localizer.FormatOrDefault("Conversion.SavedJob.Option." + name, fallback, value));
        }
        Add("Pages", "Pages: {0}", options.PageRanges);
        Add("Word", "PDF to Word: {0}", options.WordMode);
        Add("Slides", "PDF to slides: {0}", options.PowerPointMode);
        Add("Dpi", "Raster resolution: {0} DPI", options.RasterDpi);
        Add("Worksheet", "Worksheet layout: {0}", options.WorksheetLayout);
        Add("Html", "HTML layout: {0}", options.HtmlProfile);
        if (options.CompressPdfOutput) extra.Add(T("Compression", "Lossless PDF compression"));
        if (options.RequireNoLoss) extra.Add(T("StrictLoss", "Reject reported conversion loss"));
        if (extra.Count > 0) summary += Environment.NewLine + T("AdditionalOptions", "Additional retained conversion settings:") + Environment.NewLine + string.Join(Environment.NewLine, extra);
        return summary;
    }

    private string T(string name, string fallback) => _localizer.GetOrDefault("Conversion.SavedJob." + name, fallback);
}

/// <summary>One bounded source-to-output proposal shown in the folder job review.</summary>
public sealed record SavedFolderJobPlanItem(string InputName, string InputPath, string OutputPath, string Message);
