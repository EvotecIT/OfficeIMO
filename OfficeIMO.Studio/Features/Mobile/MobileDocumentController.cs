using System.ComponentModel;
using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Owns the mobile working copy and connects the shared workspace to platform file and share surfaces.</summary>
internal sealed partial class MobileDocumentController : IDisposable {
    private readonly StudioApplicationServices _services;
    private readonly string _documentsRoot;
    private readonly StudioLocalDocumentRoot _locations;
    private readonly Func<CancellationToken, Task<IStorageFile?>> _pickPdf;
    private readonly Func<string, Task> _share;
    private readonly MobileDocumentHost? _host;
    private MobileDocumentHost Host => _host ?? throw new InvalidOperationException("This document has no mobile presentation host.");
    private bool _restoring;
    private bool _sharing;
    private bool _creatingSample;
    private bool _disposed;
    private readonly HashSet<MainWindowViewModel> _observed = [];

    internal MobileDocumentController(StudioApplicationServices services,
        Func<CancellationToken, Task<IStorageFile?>> pickPdf, Func<string, Task> share, MobileDocumentHost? host = null) {
        _services = services;
        _locations = services.LocalDocuments ?? throw new ArgumentException("Mobile services require an app-owned document root.", nameof(services));
        _documentsRoot = _locations.Path;
        _pickPdf = pickPdf;
        _share = share;
        _host = host;
        Tabs = new StudioDocumentTabHost(CreateDocument, _ => {
            ActiveDocumentChanged?.Invoke(this, EventArgs.Empty);
            PersistSession();
        }, document => Host.ShowAsync<bool>(new ActiveOperationsDialogContent([document], _services.Localizer)));
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
    internal bool IsWorkingCopy => Document.DocumentPath is { } path && _locations.Resolve(_locations.GetIdentity(path)) is not null;
    internal event EventHandler? ActiveDocumentChanged;
    internal Func<MainWindowViewModel, Task<UnsavedChangesDecision>>? ConfirmUnsavedChangesAsync { get; set; }

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

    private Task<string?> PickWorkingCopyAsync(MainWindowViewModel owner, CancellationToken token) {
        if (_disposed || _creatingSample || _sharing || _restoring) return Task.FromResult<string?>(null);
        return owner.RunFileImportAsync(async cancellation => {
            IStorageFile? source = await _pickPdf(cancellation);
            return source is null ? null : await ImportAsync(source, cancellation);
        }, _locations.DiscardWorkingCopy, token);
    }

    /// <summary>Reads through the provider's permission-scoped stream and retains a separate local working copy.</summary>
    internal async Task<string> ImportAsync(IStorageFile source, CancellationToken token) {
        using var storage = new StudioStorageAccess();
        string location = await storage.RegisterAsync(source, token);
        string name = Path.GetFileName(source.Name);
        if (!string.Equals(Path.GetExtension(name), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Choose a PDF document.");
        StudioStorageSnapshot snapshot = await storage.ReadSnapshotAsync(location, token, StudioPdfSecurityPolicy.MaximumInputBytes);
        // Encrypted files defer content inspection to the shared workspace's password prompt.
        await Task.Run(() => {
            try { OfficeIMO.Pdf.PdfDocument.Load(snapshot.Bytes, StudioPdfSecurityPolicy.CreateLoadOptions()).InspectForViewing(cancellationToken: token); }
            catch (OfficeIMO.Pdf.PdfPasswordRequiredException) { }
        }, token);
        return await _locations.WriteWorkingCopyAsync(name, snapshot.Bytes, token);
    }

    internal async Task RestoreAsync(CancellationToken token = default) {
        if (!_services.Preferences.Current.RememberSession) return;
        StudioSessionSnapshot snapshot = _services.DocumentHistory.RestartSession.Load();
        _restoring = true;
        try {
            foreach (StudioSessionDocument previous in snapshot.Documents) {
                string? path = _locations.Resolve(previous.Path);
                if (path is not null) {
                    if (!File.Exists(path)) continue;
                } else if (previous.Storage is { } reference) {
                    _services.Storage.Remember(reference);
                    path = previous.Path;
                } else continue;
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
            path = document.DocumentPath ?? throw new IOException("The saved document has no location to share.");
            PersistSession();
            await MobileFileSharing.ShareAsync(_services, path, _share);
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
                .Select(state => {
                    string identity = _locations.GetIdentity(state.Path);
                    return state with { Path = identity, Storage = _locations.Resolve(identity) is not null ? null : state.Storage };
                }).ToArray();
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
