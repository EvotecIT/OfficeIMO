using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// One open document behind a native handle. It wraps the same workspace the desktop Studio uses, so editing, undo,
/// identity checks and verified saving behave identically in every host. Calls on one document are serialized; the
/// host may call from any thread and should not call from its main thread.
/// </summary>
internal sealed class StudioNativeDocument : IDisposable {
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly PdfWorkspace _workspace;
    private PdfDocumentSession? _session;
    private long _sessionRevision = -1;
    private bool _disposed;

    private StudioNativeDocument(PdfWorkspace workspace) => _workspace = workspace;

    /// <summary>
    /// Opens a PDF the host can currently read. On Apple platforms the host resolves the security-scoped URL and keeps
    /// access open for the duration of each call that reads or writes the file.
    /// </summary>
    internal static async Task<StudioNativeDocument> OpenAsync(string path, string? password, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A document path is required.", nameof(path));
        PdfWorkspace workspace = await PdfWorkspace.OpenAsync(path, cancellationToken, StudioNativeEngine.CreateRecoveryStore(), password).ConfigureAwait(false);
        return new StudioNativeDocument(workspace);
    }

    internal async Task<StudioNativeDocumentInfo> GetInfoAsync(CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            PdfDocumentViewInfo view = _workspace.ViewInfo;
            return new StudioNativeDocumentInfo(
                StudioNativeAbi.Version, _workspace.FileName, _workspace.Path, _workspace.FileSize, _workspace.Revision,
                _workspace.IsDirty, _workspace.CanUndo, _workspace.CanRedo, _workspace.RecoveryPath is not null, view.CanExtractText, view.CanExtractContent,
                _workspace.Pages.Select(page => new StudioNativePageInfo(page.PageNumber, page.Width, page.Height, page.RotationDegrees)).ToArray(),
                StudioPdfSecurityPolicy.MaximumRasterPixels);
        } finally { _gate.Release(); }
    }

    /// <summary>
    /// Renders one page as PNG. Scale is pixels per PDF point; the host applies its display scale and keeps the page
    /// within <see cref="StudioPdfSecurityPolicy.MaximumRasterPixels"/>, refused otherwise as an invalid argument.
    /// </summary>
    internal async Task<PdfRenderedPage> RenderPageAsync(int pageNumber, double scale, CancellationToken cancellationToken) {
        if (!double.IsFinite(scale) || scale <= 0D) throw new ArgumentOutOfRangeException(nameof(scale), "The scale must be a positive number.");
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            // A reader session is a snapshot of one revision; edits create a new one on the next render.
            if (_session is null || _sessionRevision != _workspace.Revision) {
                _session = PdfDocumentSession.FromWorkspace(_workspace);
                _sessionRevision = _workspace.Revision;
            }
            return await _session.RenderPageAsync(pageNumber, scale, cancellationToken).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    internal Task RotatePagesAsync(IReadOnlyList<int> pageNumbers, int degrees, CancellationToken cancellationToken) {
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        if (degrees % 90 != 0) throw new ArgumentOutOfRangeException(nameof(degrees), "Pages rotate in quarter turns.");
        return RunAsync(() => _workspace.RotateAsync(pageNumbers, degrees, cancellationToken), cancellationToken);
    }

    internal Task UndoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.UndoAsync(cancellationToken), cancellationToken);

    internal Task RedoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.RedoAsync(cancellationToken), cancellationToken);

    /// <summary>Saves in place when <paramref name="path"/> is null, otherwise to that location, which becomes the document's location.</summary>
    internal Task SaveAsync(string? path, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.SaveAsync(path, cancellationToken), cancellationToken);

    private async Task RunAsync(Func<Task> operation, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try { await operation().ConfigureAwait(false); }
        finally { _gate.Release(); }
    }

    private async Task EnterAsync(CancellationToken cancellationToken) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        if (_disposed) {
            _gate.Release();
            throw new ObjectDisposedException(nameof(StudioNativeDocument));
        }
    }

    public void Dispose() {
        if (_disposed) return;
        // Wait for a running call so the workspace is never released underneath it.
        _gate.Wait();
        try {
            _disposed = true;
            _workspace.Dispose();
        } finally { _gate.Release(); }
    }
}

internal sealed record StudioNativeDocumentInfo(
    int AbiVersion, string FileName, string Path, long FileSize, long Revision, bool IsDirty, bool CanUndo, bool CanRedo, bool HasRecovery,
    bool CanExtractText, bool CanExtractContent, IReadOnlyList<StudioNativePageInfo> Pages, long MaxRenderPixels);

internal sealed record StudioNativePageInfo(int Number, double Width, double Height, int Rotation);
