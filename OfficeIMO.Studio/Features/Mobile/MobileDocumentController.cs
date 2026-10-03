using System.ComponentModel;
using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Owns the mobile working copy and connects the shared workspace to platform file and share surfaces.</summary>
internal sealed class MobileDocumentController : IDisposable {
    private readonly StudioApplicationServices _services;
    private readonly string _documentsRoot;
    private readonly StudioLocalDocumentRoot _locations;
    private readonly Func<CancellationToken, Task<IStorageFile?>> _pickPdf;
    private readonly Func<string, Task> _share;
    private bool _restoring;
    private bool _sharing;
    private bool _creatingSample;
    private bool _disposed;
    private readonly HashSet<MainWindowViewModel> _observed = [];

    internal MobileDocumentController(StudioApplicationServices services,
        Func<CancellationToken, Task<IStorageFile?>> pickPdf, Func<string, Task> share) {
        _services = services;
        _locations = services.LocalDocuments ?? throw new ArgumentException("Mobile services require an app-owned document root.", nameof(services));
        _documentsRoot = _locations.Path;
        _pickPdf = pickPdf;
        _share = share;
        Tabs = new StudioDocumentTabHost(CreateDocument, _ => {
            ActiveDocumentChanged?.Invoke(this, EventArgs.Empty);
            PersistSession();
        });
        Tabs.Tabs.CollectionChanged += (_, _) => {
            foreach (var document in _observed.Where(document => !Tabs.OperationDocuments.Contains(document)).ToArray()) {
                document.PropertyChanged -= OnDocumentChanged;
                _observed.Remove(document);
            }
            PersistSession();
        };
    }

    internal StudioDocumentTabHost Tabs { get; }
    internal MainWindowViewModel Document => Tabs.ActiveDocument;
    internal event EventHandler? ActiveDocumentChanged;

    private MainWindowViewModel CreateDocument(Func<string, CancellationToken, Task> openInTab) {
        var document = new MainWindowViewModel(PickWorkingCopyAsync, services: _services,
            openDocumentInTab: openInTab,
            confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Save));
        document.PropertyChanged += OnDocumentChanged;
        _observed.Add(document);
        return document;
    }

    internal async Task OpenSampleAsync() {
        if (_creatingSample || _sharing || _restoring || Document.IsWorkspaceBusy || Document.IsOpening || Document.OpenCommand.IsRunning) return;
        _creatingSample = true;
        try {
            string directory = Path.Combine(_documentsRoot, Guid.NewGuid().ToString("N"));
            string path = Path.Combine(directory, "Welcome to Studio.pdf");
            await Task.Run(() => {
                Directory.CreateDirectory(directory);
                MobileSampleDocument.Save(path);
            });
            await Tabs.OpenDocumentAsync(path);
        } finally { _creatingSample = false; }
    }

    private async Task<string?> PickWorkingCopyAsync(CancellationToken token) {
        if (_creatingSample || _sharing || _restoring || Document.IsWorkspaceBusy || Document.IsOpening) return null;
        try {
            IStorageFile? source = await _pickPdf(token);
            if (source is null) return null;
            return await ImportAsync(source, token);
        } catch (OperationCanceledException) when (token.IsCancellationRequested) { return null; }
        catch (Exception error) { Document.ErrorMessage = error.Message; return null; }
    }

    /// <summary>Reads through the provider's permission-scoped stream and retains a separate local working copy.</summary>
    internal async Task<string> ImportAsync(IStorageFile source, CancellationToken token) {
        using var storage = new StudioStorageAccess();
        string location = await storage.RegisterAsync(source, token);
        string name = Path.GetFileName(source.Name);
        if (!string.Equals(Path.GetExtension(name), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Choose a PDF document.");
        StudioStorageSnapshot snapshot = await storage.ReadSnapshotAsync(location, token, StudioPdfSecurityPolicy.MaximumInputBytes);
        // Validate before retaining the import. The workspace still performs its normal security/preflight checks.
        await Task.Run(() => OfficeIMO.Pdf.PdfDocument.Load(snapshot.Bytes, StudioPdfSecurityPolicy.CreateLoadOptions()).InspectForViewing(cancellationToken: token), token);
        string directory = Path.Combine(_documentsRoot, Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(directory, name);
        await Task.Run(() => {
            token.ThrowIfCancellationRequested();
            Directory.CreateDirectory(directory);
            OfficeFileCommit.WriteAllBytes(destination, snapshot.Bytes, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
        }, token);
        return destination;
    }

    internal async Task RestoreAsync(CancellationToken token = default) {
        if (!_services.Preferences.Current.RememberSession) return;
        StudioSessionSnapshot snapshot = _services.DocumentHistory.RestartSession.Load();
        _restoring = true;
        try {
            foreach (StudioSessionDocument previous in snapshot.Documents) {
                if (_locations.Resolve(previous.Path) is not { } path || !File.Exists(path)) continue;
                await Tabs.OpenDocumentAsync(path, token);
                if (Document.DocumentPath != path) continue;
                if (Document.HasRecovery) await Document.RestoreRecoveryCommand.ExecuteAsync(null);
                Document.RestoreSessionViewState(previous.View);
            }
            if (Tabs.Tabs.FirstOrDefault(tab => tab.SourcePath is { } path && _locations.GetIdentity(path) == snapshot.ActivePath) is { } active)
                Tabs.SelectedTab = active;
        } finally { _restoring = false; }
        PersistSession();
    }

    internal async Task ShareAsync() {
        if (_sharing || _creatingSample || _restoring || Document.OpenCommand.IsRunning || Document.IsWorkspaceBusy || Document.IsOpening || Document.DocumentPath is not { } path) return;
        _sharing = true;
        MainWindowViewModel document = Document;
        try {
            document.ErrorMessage = null;
            if (document.IsDirty) await document.SaveCommand.ExecuteAsync(null);
            if (document.IsDirty || document.HasError) return;
            PersistSession();
            await _share(path);
        } finally { _sharing = false; }
    }

    internal void Suspend() {
        foreach (var tab in Tabs.Tabs) tab.Document.SaveDocumentViewState();
        PersistSession();
        foreach (var document in Tabs.OperationDocuments) document.SetPresentationActive(false);
    }

    internal void Resume() {
        Document.SetPresentationActive(true);
        Document.SelectedPage?.AttachToViewport();
        ActiveDocumentChanged?.Invoke(this, EventArgs.Empty);
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName is nameof(MainWindowViewModel.IsDirty) or nameof(MainWindowViewModel.SelectedPage) ||
            args.PropertyName == nameof(MainWindowViewModel.IsOpening) && !Document.IsOpening) PersistSession();
    }

    private void PersistSession() {
        if (_disposed || _restoring || Tabs is null || !_services.Preferences.Current.RememberSession || Tabs.OperationDocuments.Any(document => document.IsOpening)) return;
        try {
            var states = Tabs.Tabs.Select(tab => tab.Document.CaptureSessionDocument())
                .OfType<StudioSessionDocument>()
                .Select(state => state with { Path = _locations.GetIdentity(state.Path), Storage = null }).ToArray();
            string? active = Document.DocumentPath is { } path ? _locations.GetIdentity(path) : null;
            _services.DocumentHistory.RestartSession.Save(new(1, DateTimeOffset.UtcNow, active, states));
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            Document.ErrorMessage = "The reading position could not be saved: " + error.Message;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        Suspend();
        _disposed = true;
        foreach (var document in _observed) document.PropertyChanged -= OnDocumentChanged;
        _observed.Clear();
        Tabs.Dispose();
    }
}
