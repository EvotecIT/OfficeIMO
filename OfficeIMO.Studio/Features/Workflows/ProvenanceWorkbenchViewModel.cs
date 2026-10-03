using System.Collections.ObjectModel;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Thin local-file workspace over the canonical provenance workflow and report.</summary>
public sealed partial class ProvenanceWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickInput;
    private readonly Func<CancellationToken, Task<string?>> _pickFolder;
    private readonly IOfficeProvenanceWorkflowRunner _runner;
    private readonly IOfficeWorkflowPublicationGuard? _publicationGuard;
    private readonly StudioJobHistory? _jobs;
    private OfficeProvenanceWorkflowResult? _review, _lastResult;
    private CancellationTokenSource? _cancellation;
    private int _revision;
    private bool _disposed;

    public ProvenanceWorkbenchViewModel(Func<CancellationToken, Task<string?>> pickInput,
        Func<CancellationToken, Task<string?>> pickFolder, IOfficeProvenanceWorkflowRunner? runner = null, IOfficeWorkflowPublicationGuard? publicationGuard = null)
        : this(pickInput, pickFolder, runner, publicationGuard, null) { }

    internal ProvenanceWorkbenchViewModel(Func<CancellationToken, Task<string?>> pickInput,
        Func<CancellationToken, Task<string?>> pickFolder, IOfficeProvenanceWorkflowRunner? runner,
        IOfficeWorkflowPublicationGuard? publicationGuard, StudioJobHistory? jobHistory) {
        _pickInput = pickInput; _pickFolder = pickFolder; _runner = runner ?? new OfficeWorkflowRunner(); _publicationGuard = publicationGuard;
        _jobs = jobHistory;
    }
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanAssess)), NotifyPropertyChangedFor(nameof(CanCreateCopy))]
    private string _inputPath = "";
    [ObservableProperty] private string _outputFolder = "";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanAssess)), NotifyPropertyChangedFor(nameof(CanCreateCopy)), NotifyPropertyChangedFor(nameof(CanExportReport))]
    private bool _isBusy;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeManifests;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeReferences;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeDeclarations;
    [ObservableProperty] private string _status = "Choose a local file and assess its supported provenance evidence.";
    [ObservableProperty] private string _checks = "Structural: NotRequested · Text integrity: NotRequested · Verification: NotConfigured · Providers: NotConfigured";
    [ObservableProperty] private string _outputPath = "";
    [ObservableProperty] private string _reportPath = "";
    [ObservableProperty] private string _inputHash = "";
    [ObservableProperty] private string _outputHash = "";
    [ObservableProperty] private string _coverage = "";
    public ObservableCollection<string> Findings { get; } = new();
    public ObservableCollection<string> Changes { get; } = new();
    public ObservableCollection<string> Diagnostics { get; } = new();
    public bool CanAssess => !_disposed && !IsBusy && !string.IsNullOrWhiteSpace(InputPath);
    public bool CanCreateCopy => CanAssess && _review?.Succeeded == true && _review.InputSha256 != null &&
        OfficeProvenanceWorkflowCatalog.FindByPath(InputPath)?.CanRemove == true && (RemoveManifests || RemoveReferences || RemoveDeclarations);
    public bool CanExportReport => !_disposed && !IsBusy && _lastResult != null;
    public bool CanCancel => IsBusy;
    partial void OnInputPathChanged(string value) {
        _revision++; _review = _lastResult = null;
        Findings.Clear(); Changes.Clear(); Diagnostics.Clear(); OutputPath = ReportPath = InputHash = OutputHash = Coverage = "";
        Checks = "Structural: NotRequested · Text integrity: NotRequested · Verification: NotConfigured · Providers: NotConfigured";
        OnPropertyChanged(nameof(CanExportReport));
    }
    partial void OnIsBusyChanged(bool value) => OnPropertyChanged(nameof(CanCancel));
    [RelayCommand] private async Task ChooseInputAsync() {
        if (IsBusy || _disposed) return;
        int revision = _revision;
        string? path = await _pickInput(CancellationToken.None);
        if (!_disposed && revision == _revision && path != null) InputPath = path;
    }
    [RelayCommand] private async Task ChooseFolderAsync() {
        if (IsBusy || _disposed) return;
        string? folder = await _pickFolder(CancellationToken.None);
        if (!_disposed && folder != null) OutputFolder = folder;
    }
    [RelayCommand] private Task AssessAsync() => RunAsync(false);
    [RelayCommand] private Task CreateCopyAsync() => RunAsync(true);
    private async Task RunAsync(bool remove) {
        if (remove ? !CanCreateCopy : !CanAssess) return;
        if (!Path.IsPathFullyQualified(InputPath) || !File.Exists(InputPath)) { Status = "Select an accessible local file. Provider locations are not supported by this workbench."; return; }
        if (remove && (!Path.IsPathFullyQualified(OutputFolder) || !Directory.Exists(OutputFolder))) {
            Status = "Choose an existing local output folder for the separate copy."; return;
        }
        int revision = _revision;
        var request = new OfficeProvenanceWorkflowRequest { InputPath = InputPath,
            Operation = remove ? OfficeProvenanceWorkflowOperation.Remove : OfficeProvenanceWorkflowOperation.Assess,
            ExpectedInputSha256 = remove ? _review!.InputSha256 : null,
            OutputPath = remove ? Path.Combine(OutputFolder, Path.GetFileNameWithoutExtension(InputPath) + ".provenance-cleaned" + Path.GetExtension(InputPath)) : null,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Rename, PublicationGuard = _publicationGuard };
        request.Removal.RemoveC2paManifests = RemoveManifests;
        request.Removal.RemoveExternalC2paReferences = RemoveReferences;
        request.Removal.RemoveAiSourceMetadata = RemoveDeclarations;
        request.Removal.SignatureMutationPolicy = OfficeSignatureMutationPolicy.BlockSave;
        IsBusy = true; Status = remove ? "Creating and re-inspecting a separate copy…" : "Assessing local file…";
        using var cancellation = new CancellationTokenSource(); _cancellation = cancellation;
        StudioJobRecord? job = null;
        bool ownerStarted = false;
        try {
            job = _jobs?.Start(remove ? "Provenance copy" : "Provenance assessment", request.InputPath, request.OutputPath, cancellation.Cancel);
            using IDisposable? permit = _jobs is null ? null : await _jobs.EnterAsync(cancellation.Token);
            job?.Report(new("provenance", "execute", Status, 0, 0));
            ownerStarted = true;
            OfficeProvenanceWorkflowResult result = await _runner.RunProvenanceAsync(request, cancellationToken: cancellation.Token);
            job?.Complete(result.Status, result.OutputPath, result.Summary);
            if (_disposed || revision != _revision) return;
            _lastResult = result;
            if (!remove) _review = result.Succeeded ? result : null;
            else if (!result.Succeeded) _review = null;
            Status = result.Summary; InputHash = result.InputSha256 ?? ""; OutputHash = result.OutputSha256 ?? "";
            OutputPath = result.OutputPath ?? ""; ReportPath = "";
            ProvenanceResultDto document = OfficeProvenanceReportSerializer.Create(result);
            Checks = $"Structural: {document.Checks.Structural} · Text integrity: {document.Checks.TextIntegrity} · Verification: {document.Checks.Verification} · Providers: {document.Checks.ProviderSignals}";
            Coverage = document.CoverageNotes;
            Findings.Clear(); Changes.Clear(); Diagnostics.Clear();
            OfficeProvenanceReport? report = result.Assessment?.Structural ?? result.Inspection ?? result.After;
            foreach (OfficeProvenanceEvidence evidence in report?.Evidence ?? Array.Empty<OfficeProvenanceEvidence>())
                Findings.Add($"{evidence.Carrier} · {evidence.Location} · {(evidence.IsStructurallyValid ? "Structurally recognized" : "Malformed or ambiguous")}");
            foreach (OfficeTextIntegrityFinding finding in result.Assessment?.TextIntegrity?.Findings ?? Array.Empty<OfficeTextIntegrityFinding>())
                Findings.Add($"{finding.UnicodeNotation} · {finding.Kind} · {finding.Risk} · UTF-16 offset {finding.TextOffset}");
            foreach (OfficeProvenanceChange change in result.Changes) Changes.Add($"{change.Carrier} · {change.Location} · {change.RemovedBytes} bytes removed");
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) Diagnostics.Add($"{diagnostic.Severity}: {diagnostic.Message}");
            foreach (string diagnostic in report?.Diagnostics ?? Array.Empty<string>()) Diagnostics.Add(diagnostic);
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            string message = remove && ownerStarted ? "Cancelled. Check the output folder before retrying." : "Cancelled.";
            job?.Complete(remove && ownerStarted ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Cancelled, null, message);
            if (!_disposed) Status = message;
        }
        catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string message = "Provenance workflow failed: " + error.Message;
            job?.Complete(remove && ownerStarted ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Failed, null, message);
            if (!_disposed) Status = message;
        }
        finally { _cancellation = null; IsBusy = false; OnPropertyChanged(nameof(CanCreateCopy)); OnPropertyChanged(nameof(CanExportReport)); }
    }
    [RelayCommand] private async Task ExportReportAsync() {
        if (!CanExportReport) return;
        if (!Path.IsPathFullyQualified(OutputFolder) || !Directory.Exists(OutputFolder)) { Status = "Choose an existing local folder for the report."; return; }
        var result = _lastResult!; int revision = _revision;
        string path = Path.Combine(OutputFolder, "provenance-report-" + Guid.NewGuid().ToString("N") + ".json");
        using var cancellation = new CancellationTokenSource(); _cancellation = cancellation;
        StudioJobRecord? job = null;
        IsBusy = true;
        async Task VerifyDestinationAsync(CancellationToken token) {
            if (_disposed || revision != _revision) cancellation.Cancel();
            token.ThrowIfCancellationRequested();
            if (_publicationGuard != null && !await _publicationGuard.CanPublishAsync(path, false, token).ConfigureAwait(false))
                throw new IOException("The report destination is protected by the host publication policy.");
            token.ThrowIfCancellationRequested();
        }
        try {
            job = _jobs?.Start("Provenance report", result.InputPath ?? InputPath, path, cancellation.Cancel);
            Status = "Queued";
            using IDisposable? permit = _jobs is null ? null : await _jobs.EnterAsync(cancellation.Token);
            Status = "Exporting report…";
            job?.Report(new("provenance-report", "publish", Status, 0, 0));
            await VerifyDestinationAsync(cancellation.Token);
            byte[] bytes = Encoding.UTF8.GetBytes(OfficeProvenanceReportSerializer.Serialize(result));
            await OfficeFileCommit.WriteAsync(path, async (stream, token) => {
                await stream.WriteAsync(bytes, token).ConfigureAwait(false);
                await VerifyDestinationAsync(token).ConfigureAwait(false);
            }, OfficeFileCommit.ConflictPolicy.FailIfExists, cancellation.Token);
            job?.Complete(OfficeWorkflowStatus.Completed, path, "Report exported.");
            if (_disposed || revision != _revision) return;
            ReportPath = path; Status = "Report exported.";
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            job?.Complete(OfficeWorkflowStatus.Cancelled, null, "Report export cancelled.");
            if (!_disposed) Status = "Report export cancelled.";
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string message = "Report export failed: " + error.Message;
            job?.Complete(OfficeWorkflowStatus.Failed, null, message);
            if (!_disposed) Status = message;
        }
        finally { _cancellation = null; IsBusy = false; }
    }
    [RelayCommand] private void Cancel() => _cancellation?.Cancel();
    public void Dispose() { if (_disposed) return; _disposed = true; _revision++; _cancellation?.Cancel(); }
}
