using OfficeIMO.Studio.Infrastructure.Localization;
using System.Collections.ObjectModel;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;
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
    /// <summary>Whether the host imports private copies and shares results instead of exposing local paths.</summary>
    public bool UsesWorkingCopies { get; internal set; }
    public string InputName => Path.GetFileName(InputPath);
    public string OutputName => Path.GetFileName(OutputPath);
    public string ReportName => Path.GetFileName(ReportPath);
    /// <summary>The copy and report as shown: file names for private working copies, full paths otherwise.</summary>
    public string CopyDisplay => UsesWorkingCopies ? OutputName : OutputPath;
    public string ReportDisplay => UsesWorkingCopies ? ReportName : ReportPath;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(InputName)), NotifyPropertyChangedFor(nameof(CanAssess)), NotifyPropertyChangedFor(nameof(CanCreateCopy))]
    private string _inputPath = "";
    [ObservableProperty] private string _outputFolder = "";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanAssess)), NotifyPropertyChangedFor(nameof(CanCreateCopy)), NotifyPropertyChangedFor(nameof(CanExportReport))]
    private bool _isBusy;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeManifests;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeReferences;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanCreateCopy))] private bool _removeDeclarations;
    [ObservableProperty] private string _status = L("Provenance.Status.Ready");
    [ObservableProperty] private string _checks = EmptyChecks();
    [ObservableProperty, NotifyPropertyChangedFor(nameof(OutputName)), NotifyPropertyChangedFor(nameof(CopyDisplay))] private string _outputPath = "";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(ReportName)), NotifyPropertyChangedFor(nameof(ReportDisplay))] private string _reportPath = "";
    [ObservableProperty] private string _inputHash = "";
    [ObservableProperty] private string _outputHash = "";
    [ObservableProperty] private string _coverage = "";
    public ObservableCollection<string> Findings { get; } = new();
    public ObservableCollection<string> Changes { get; } = new();
    public ObservableCollection<string> Diagnostics { get; } = new();
    /// <summary>Active credential statements for the current assessment or resulting copy.</summary>
    public ObservableCollection<ProvenanceCredentialViewModel> Credentials { get; } = new();
    /// <summary>Whether the current report contains a readable credential summary.</summary>
    public bool HasCredentials => Credentials.Count > 0;
    public bool CanAssess => !_disposed && !IsBusy && !string.IsNullOrWhiteSpace(InputPath);
    public bool CanCreateCopy => CanAssess && _review?.Succeeded == true && _review.InputSha256 != null &&
        OfficeProvenanceWorkflowCatalog.FindByPath(InputPath)?.CanRemove == true && (RemoveManifests || RemoveReferences || RemoveDeclarations);
    public bool CanExportReport => !_disposed && !IsBusy && _lastResult != null;
    public bool CanCancel => IsBusy;
    partial void OnInputPathChanged(string value) {
        _revision++; _review = _lastResult = null;
        Findings.Clear(); Changes.Clear(); Diagnostics.Clear(); OutputPath = ReportPath = InputHash = OutputHash = Coverage = "";
        ClearCredentials();
        Checks = EmptyChecks();
        OnPropertyChanged(nameof(CanExportReport));
    }
    partial void OnIsBusyChanged(bool value) => OnPropertyChanged(nameof(CanCancel));
    [RelayCommand] private async Task ChooseInputAsync() {
        if (IsBusy || _disposed) return;
        int revision = _revision;
        using var cancellation = new CancellationTokenSource();
        _cancellation = cancellation;
        IsBusy = true;
        try {
            string? path = await _pickInput(cancellation.Token);
            cancellation.Token.ThrowIfCancellationRequested();
            if (!_disposed && revision == _revision && path != null) InputPath = path;
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            if (!_disposed) Status = L("Provenance.Status.ImportCancelled");
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status = error.Message; }
        finally { _cancellation = null; IsBusy = false; }
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
        if (!Path.IsPathFullyQualified(InputPath) || !File.Exists(InputPath)) { Status = L("Provenance.Status.SelectFile"); return; }
        if (remove && (!Path.IsPathFullyQualified(OutputFolder) || !Directory.Exists(OutputFolder))) {
            Status = L("Provenance.Status.ChooseOutputFolder"); return;
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
        _lastResult = null;
        if (!remove) _review = null;
        Findings.Clear(); Changes.Clear(); Diagnostics.Clear(); ClearCredentials();
        OutputPath = ReportPath = InputHash = OutputHash = Coverage = "";
        Checks = "Structural: NotRequested · Text integrity: NotRequested · Verification: NotConfigured · Providers: NotConfigured";
        IsBusy = true; Status = L(remove ? "Provenance.Status.Copying" : "Provenance.Status.Assessing");
        using var cancellation = new CancellationTokenSource(); _cancellation = cancellation;
        StudioJobRecord? job = null;
        bool ownerStarted = false;
        try {
            job = _jobs?.Start(L(remove ? "Provenance.Job.Copy" : "Provenance.Job.Assessment"), request.InputPath, request.OutputPath, cancellation.Cancel);
            using IDisposable? permit = _jobs is null ? null : await _jobs.EnterAsync(cancellation.Token);
            job?.Report(new("provenance", "execute", Status, 0, 0));
            ownerStarted = true;
            IOfficeProvenanceWorkflowRunner runner = _runner;
            if (!remove && CanConfigureProvider && !string.IsNullOrWhiteSpace(C2paToolPath)) {
                runner = new OfficeWorkflowRunner(new C2paToolProvenanceVerifier(C2paToolPath.Trim()));
                request.Assessment.Verification.TrustAnchorsPath = string.IsNullOrWhiteSpace(TrustAnchorsPath) ? null : TrustAnchorsPath.Trim();
                request.Assessment.Verification.AllowedListPath = string.IsNullOrWhiteSpace(AllowedListPath) ? null : AllowedListPath.Trim();
            }
            OfficeProvenanceWorkflowResult result = await runner.RunProvenanceAsync(request, cancellationToken: cancellation.Token);
            job?.Complete(result.Status, result.OutputPath, result.Summary);
            if (_disposed || revision != _revision) return;
            _lastResult = result;
            if (!remove) _review = result.Succeeded ? result : null;
            else if (!result.Succeeded) _review = null;
            Status = result.Summary; InputHash = result.InputSha256 ?? ""; OutputHash = result.OutputSha256 ?? "";
            OutputPath = result.OutputPath ?? ""; ReportPath = "";
            ProvenanceResultDto document = OfficeProvenanceReportSerializer.Create(result);
            Checks = F("Provenance.Checks", document.Checks.Structural, document.Checks.TextIntegrity, document.Checks.Verification, document.Checks.ProviderSignals);
            Coverage = document.CoverageNotes;
            Findings.Clear(); Changes.Clear(); Diagnostics.Clear();
            OfficeProvenanceReport? report = result.Assessment?.Structural ?? result.Inspection ?? result.After;
            ProvenanceReportDto? transport = document.Assessment?.Structural ?? document.Inspection ?? document.After;
            foreach (ProvenanceEvidenceDto evidence in transport?.Evidence ?? Array.Empty<ProvenanceEvidenceDto>()) {
                if (evidence.Manifest is { } manifest) Credentials.Add(new ProvenanceCredentialViewModel(evidence.Location, manifest));
            }
            OnPropertyChanged(nameof(HasCredentials));
            foreach (OfficeProvenanceEvidence evidence in report?.Evidence ?? Array.Empty<OfficeProvenanceEvidence>())
                Findings.Add(F(evidence.IsStructurallyValid ? "Provenance.Evidence.Recognized" : "Provenance.Evidence.Malformed", evidence.Carrier, evidence.Location));
            foreach (OfficeTextIntegrityFinding finding in result.Assessment?.TextIntegrity?.Findings ?? Array.Empty<OfficeTextIntegrityFinding>())
                Findings.Add(F("Provenance.TextFinding", finding.UnicodeNotation, finding.Kind, finding.Risk, finding.Location, finding.TextOffset));
            if (result.Assessment?.Verification is { } verification) {
                Findings.Add(F("Provenance.VerificationFinding", verification.ProviderName, verification.Status));
                foreach (string finding in verification.Findings) Findings.Add(finding);
            }
            foreach (var signal in result.Assessment?.ProviderSignals ?? Array.Empty<OfficeProvenanceSignalResult>()) {
                Findings.Add($"{signal.ProviderName} · {signal.SignalKind} · {signal.Status}");
                if (signal.Measurement is { } measurement)
                    Findings.Add(F("Provenance.Measurement", measurement.Algorithm, measurement.DetectorVersion, measurement.ScoreName, measurement.Score, measurement.Threshold, measurement.TokenCount));
                foreach (string finding in signal.Findings) Findings.Add(finding);
            }
            foreach (OfficeProvenanceChange change in result.Changes) Changes.Add(F("Provenance.Change", change.Carrier, change.Location, change.RemovedBytes));
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) Diagnostics.Add($"{diagnostic.Severity}: {diagnostic.Message}");
            foreach (string diagnostic in report?.Diagnostics ?? Array.Empty<string>()) Diagnostics.Add(diagnostic);
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            string message = L(remove && ownerStarted ? "Provenance.Status.CancelledCheckOutput" : "Provenance.Status.Cancelled");
            job?.Complete(remove && ownerStarted ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Cancelled, null, message);
            if (!_disposed) { Status = message; _review = null; }
        }
        catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string message = F("Provenance.Status.Failed", error.Message);
            job?.Complete(remove && ownerStarted ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Failed, null, message);
            if (!_disposed) { Status = message; _review = null; }
        }
        finally { _cancellation = null; IsBusy = false; OnPropertyChanged(nameof(CanCreateCopy)); OnPropertyChanged(nameof(CanExportReport)); }
    }
    [RelayCommand] private async Task ExportReportAsync() {
        if (!CanExportReport) return;
        if (!Path.IsPathFullyQualified(OutputFolder) || !Directory.Exists(OutputFolder)) { Status = L("Provenance.Status.ChooseReportFolder"); return; }
        var result = _lastResult!; int revision = _revision;
        string path = Path.Combine(OutputFolder, "provenance-report-" + Guid.NewGuid().ToString("N") + ".json");
        using var cancellation = new CancellationTokenSource(); _cancellation = cancellation;
        StudioJobRecord? job = null;
        IsBusy = true;
        async Task VerifyDestinationAsync(CancellationToken token) {
            if (_disposed || revision != _revision) cancellation.Cancel();
            token.ThrowIfCancellationRequested();
            if (_publicationGuard != null && !await _publicationGuard.CanPublishAsync(path, false, token).ConfigureAwait(false))
                throw new IOException(L("Provenance.Status.ReportProtected"));
            token.ThrowIfCancellationRequested();
        }
        try {
            job = _jobs?.Start(L("Provenance.Job.Report"), result.InputPath ?? InputPath, path, cancellation.Cancel);
            Status = L("Provenance.Status.Queued");
            using IDisposable? permit = _jobs is null ? null : await _jobs.EnterAsync(cancellation.Token);
            Status = L("Provenance.Status.Exporting");
            job?.Report(new("provenance-report", "publish", Status, 0, 0));
            await VerifyDestinationAsync(cancellation.Token);
            byte[] bytes = Encoding.UTF8.GetBytes(OfficeProvenanceReportSerializer.Serialize(result));
            await OfficeFileCommit.WriteAsync(path, async (stream, token) => {
                await stream.WriteAsync(bytes, token).ConfigureAwait(false);
                await VerifyDestinationAsync(token).ConfigureAwait(false);
            }, OfficeFileCommit.ConflictPolicy.FailIfExists, cancellation.Token);
            job?.Complete(OfficeWorkflowStatus.Completed, path, L("Provenance.Status.ReportExported"));
            if (_disposed || revision != _revision) return;
            ReportPath = path; Status = L("Provenance.Status.ReportExported");
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            job?.Complete(OfficeWorkflowStatus.Cancelled, null, L("Provenance.Status.ReportCancelled"));
            if (!_disposed) Status = L("Provenance.Status.ReportCancelled");
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string message = F("Provenance.Status.ReportFailed", error.Message);
            job?.Complete(OfficeWorkflowStatus.Failed, null, message);
            if (!_disposed) Status = message;
        }
        finally { _cancellation = null; IsBusy = false; }
    }
    [RelayCommand] private void Cancel() => _cancellation?.Cancel();
    private static string L(string key) => StudioLocalization.Current.Get(key);
    private static string F(string key, params object?[] arguments) => StudioLocalization.Current.Format(key, arguments);
    private static string EmptyChecks() => F("Provenance.Checks", "NotRequested", "NotRequested", "NotConfigured", "NotConfigured");
    private void ClearCredentials() { Credentials.Clear(); OnPropertyChanged(nameof(HasCredentials)); }
    public void Dispose() { if (_disposed) return; _disposed = true; _revision++; _cancellation?.Cancel(); }
}
