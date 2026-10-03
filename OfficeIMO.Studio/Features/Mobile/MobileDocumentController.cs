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

    internal MobileDocumentController(StudioApplicationServices services,
        Func<CancellationToken, Task<IStorageFile?>> pickPdf, Func<string, Task> share) {
        _services = services;
        _locations = services.LocalDocuments ?? throw new ArgumentException("Mobile services require an app-owned document root.", nameof(services));
        _documentsRoot = _locations.Path;
        _pickPdf = pickPdf;
        _share = share;
        Document = new MainWindowViewModel(PickWorkingCopyAsync, services: services,
            confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Save));
        Document.PropertyChanged += OnDocumentChanged;
    }

    internal MainWindowViewModel Document { get; }

    internal async Task OpenSampleAsync() {
        if (_creatingSample || _sharing || _restoring || Document.IsWorkspaceBusy || Document.IsOpening || Document.OpenCommand.IsRunning) return;
        _creatingSample = true;
        try {
            string directory = Path.Combine(_documentsRoot, Guid.NewGuid().ToString("N"));
            string path = Path.Combine(directory, "Welcome to Studio.pdf");
            await Task.Run(() => {
                Directory.CreateDirectory(directory);
                OfficeIMO.Pdf.PdfDocument.Create(pdf => pdf.Content(content => content
                    .H1("A little room to think.")
                    .Paragraph(p => p.Text("Welcome to OfficeIMO Studio"))
                    .H2("Read at your own pace")
                    .Paragraph(p => p.Text("Use the zoom controls to get closer. Pages opens the page list. On a wide iPad, your pages stay beside the document."))
                    .H2("Leave a thought")
                    .Paragraph(p => p.Text("Choose Note to add a review note. Undo lets you change your mind. Search finds words in the document."))
                    .H2("Make it yours")
                    .Paragraph(p => p.Text("Open a PDF from Files to create a working copy. Your original stays where it is. Share saves your changes and opens the Apple share sheet."))
                    .H2("Keep going")
                    .Paragraph(p => p.Text("Studio keeps recovery copies of edits and restores your reading position when you return."))),
                    new OfficeIMO.Pdf.PdfOptions { DefaultFont = OfficeIMO.Pdf.PdfStandardFont.Helvetica, DefaultFontSize = 13 })
                    .Meta(title: "Welcome to Studio", author: "OfficeIMO").Save(path);
            });
            await Document.OpenDocumentAsync(path);
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
        StudioSessionDocument? previous = snapshot.Documents.FirstOrDefault(item => item.Path == snapshot.ActivePath);
        if (previous is null || _locations.Resolve(previous.Path) is not { } path || !File.Exists(path)) return;
        _restoring = true;
        try {
            await Document.OpenDocumentAsync(path, token);
            if (!Document.HasDocument) return;
            if (Document.HasRecovery) await Document.RestoreRecoveryCommand.ExecuteAsync(null);
            Document.RestoreSessionViewState(previous.View);
        } finally { _restoring = false; }
        PersistSession();
    }

    internal async Task ShareAsync() {
        if (_sharing || _creatingSample || _restoring || Document.OpenCommand.IsRunning || Document.IsWorkspaceBusy || Document.IsOpening || Document.DocumentPath is not { } path) return;
        _sharing = true;
        try {
            Document.ErrorMessage = null;
            if (Document.IsDirty) await Document.SaveCommand.ExecuteAsync(null);
            if (Document.IsDirty || Document.HasError) return;
            PersistSession();
            await _share(path);
        } finally { _sharing = false; }
    }

    internal void Suspend() {
        Document.SaveDocumentViewState();
        PersistSession();
        Document.SetPresentationActive(false);
    }

    internal void Resume() {
        Document.SetPresentationActive(true);
        Document.SelectedPage?.AttachToViewport();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName is nameof(MainWindowViewModel.IsDirty) or nameof(MainWindowViewModel.SelectedPage) ||
            args.PropertyName == nameof(MainWindowViewModel.IsOpening) && !Document.IsOpening) PersistSession();
    }

    private void PersistSession() {
        if (_disposed || _restoring || !_services.Preferences.Current.RememberSession || Document.IsOpening ||
            Document.CaptureSessionDocument() is not { } state) return;
        try {
            state = state with { Path = _locations.GetIdentity(state.Path), Storage = null };
            _services.DocumentHistory.RestartSession.Save(new(1, DateTimeOffset.UtcNow, state.Path, [state]));
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            Document.ErrorMessage = "The reading position could not be saved: " + error.Message;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        Suspend();
        _disposed = true;
        Document.PropertyChanged -= OnDocumentChanged;
        Document.Dispose();
    }
}
