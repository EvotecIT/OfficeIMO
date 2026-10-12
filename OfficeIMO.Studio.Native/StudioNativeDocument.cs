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
                _workspace.Pages.Select(page => {
                    (double width, double height) = PdfDocumentSession.GetPhysicalSize(page);
                    return new StudioNativePageInfo(page.PageNumber, width, height, page.RotationDegrees);
                }).ToArray(),
                StudioPdfSecurityPolicy.MaximumRasterPixels, _workspace.CanMutatePages, _workspace.CanExtractPages);
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
            return await CurrentSession().RenderPageAsync(pageNumber, scale, cancellationToken).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    /// <summary>Renders a rectangle of a page as opaque RGBA pixels; see <see cref="PdfDocumentSession.RenderRegionAsync"/>.</summary>
    internal async Task<PdfRenderedPixels> RenderRegionAsync(int pageNumber, double scale, int x, int y, int width, int height,
        CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            return await CurrentSession().RenderRegionAsync(pageNumber, scale, x, y, width, height, cancellationToken).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    internal Task RotatePagesAsync(IReadOnlyList<int> pageNumbers, int degrees, CancellationToken cancellationToken) {
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        if (degrees % 90 != 0) throw new ArgumentOutOfRangeException(nameof(degrees), "Pages rotate in quarter turns.");
        return RunAsync(() => _workspace.RotateAsync(pageNumbers, degrees, cancellationToken), cancellationToken);
    }

    /// <summary>Deletes pages as one undoable change. A document keeps at least one page.</summary>
    internal Task DeletePagesAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken) =>
        RunAsync(() => {
            int[] pages = RequirePages(pageNumbers);
            if (pages.Length >= _workspace.Pages.Count) throw new ArgumentException("A document keeps at least one page.", nameof(pageNumbers));
            return _workspace.DeleteAsync(pages, cancellationToken);
        }, cancellationToken);

    /// <summary>Duplicates pages as one undoable change; each copy follows its page.</summary>
    internal Task DuplicatePagesAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.DuplicateAsync(RequirePages(pageNumbers), cancellationToken), cancellationToken);

    /// <summary>Puts the pages in a new order as one undoable change: <paramref name="order"/> lists every current page once.</summary>
    internal Task ReorderPagesAsync(IReadOnlyList<int> order, CancellationToken cancellationToken) =>
        RunAsync(() => {
            int count = _workspace.Pages.Count;
            if (order.Count != count || order.Distinct().Count() != count || order.Any(page => page < 1 || page > count))
                throw new ArgumentException("The new order must list every page exactly once.", nameof(order));
            return _workspace.ReorderAsync(order, cancellationToken);
        }, cancellationToken);

    /// <summary>Inserts a blank page of the given size in points before a page, or after the last page with count + 1.</summary>
    internal Task InsertBlankPageAsync(int insertBefore, double width, double height, CancellationToken cancellationToken) {
        if (!double.IsFinite(width) || !double.IsFinite(height) || width < 3D || height < 3D || width > 14400D || height > 14400D)
            throw new ArgumentOutOfRangeException(nameof(width), "A page is between 3 and 14,400 points on each side.");
        return RunAsync(() => {
            if (insertBefore < 1 || insertBefore > _workspace.Pages.Count + 1) throw new ArgumentOutOfRangeException(nameof(insertBefore), "The position is outside the document.");
            return _workspace.InsertBlankAsync(insertBefore, width, height, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>Writes pages to a new PDF at <paramref name="path"/>; the open document is unchanged.</summary>
    internal Task ExtractPagesAsync(IReadOnlyList<int> pageNumbers, string path, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination is required.", nameof(path));
        // The local path keeps the workflow runner, and the conversions it roots, out of the NativeAOT library.
        return RunAsync(() => _workspace.ExtractToLocalFileAsync(RequirePages(pageNumbers), path, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Finds text, ignoring case. Each hit's rectangles are in points of the page as displayed: from its top-left corner,
    /// after rotation and user unit, the same space as the page sizes the host lays out.
    /// </summary>
    internal async Task<StudioNativeSearchResult> SearchAsync(string query, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            PdfDocumentSession session = CurrentSession();
            if (!session.CanSearch) throw new NotSupportedException("Text search is restricted by this document's permissions.");
            IReadOnlyList<PdfSearchHit> hits = await session.SearchAsync(query, cancellationToken).ConfigureAwait(false);
            return new StudioNativeSearchResult(StudioNativeAbi.Version, query.Trim(), _workspace.Revision, hits.Select(hit => new StudioNativeSearchHit(
                hit.PageNumber, hit.OccurrenceNumber, hit.Snippet, StudioNativeRectangle.From(hit.Bounds),
                hit.LineBounds.Select(StudioNativeRectangle.From).ToArray())).ToArray());
        } finally { _gate.Release(); }
    }

    internal Task UndoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.UndoAsync(cancellationToken), cancellationToken);

    internal Task RedoAsync(CancellationToken cancellationToken) => RunAsync(() => _workspace.RedoAsync(cancellationToken), cancellationToken);

    /// <summary>Saves in place when <paramref name="path"/> is null, otherwise to that location, which becomes the document's location.</summary>
    internal Task SaveAsync(string? path, CancellationToken cancellationToken) =>
        RunAsync(() => _workspace.SaveAsync(path, cancellationToken), cancellationToken);

    // A reader session is a snapshot of one revision, with its rendering caches; edits create a new one on the next render.
    private PdfDocumentSession CurrentSession() {
        if (_session is null || _sessionRevision != _workspace.Revision) {
            _session = PdfDocumentSession.FromWorkspace(_workspace);
            _sessionRevision = _workspace.Revision;
        }
        return _session;
    }

    private int[] RequirePages(IReadOnlyList<int> pageNumbers) {
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        int count = _workspace.Pages.Count;
        if (pageNumbers.Any(page => page < 1 || page > count)) throw new ArgumentOutOfRangeException(nameof(pageNumbers), "The page does not exist.");
        return pageNumbers.Distinct().Order().ToArray();
    }

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
            // Closing discards unsaved changes, so their recovery copy goes too. A host that ends without closing keeps
            // it for the next launch, and so does a failed deletion, which close cannot report.
            if (_workspace.IsDirty) {
                try { _workspace.DiscardRecoveryAsync().GetAwaiter().GetResult(); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            }
            _workspace.Dispose();
        } finally { _gate.Release(); }
    }
}

internal sealed record StudioNativeDocumentInfo(
    int AbiVersion, string FileName, string Path, long FileSize, long Revision, bool IsDirty, bool CanUndo, bool CanRedo, bool HasRecovery,
    bool CanExtractText, bool CanExtractContent, IReadOnlyList<StudioNativePageInfo> Pages, long MaxRenderPixels, bool CanMutatePages,
    bool CanExtractPages);

internal sealed record StudioNativePageInfo(int Number, double Width, double Height, int Rotation);

internal sealed record StudioNativeSearchResult(int AbiVersion, string Query, long Revision, IReadOnlyList<StudioNativeSearchHit> Hits);

internal sealed record StudioNativeSearchHit(int Page, int Occurrence, string Text, StudioNativeRectangle Bounds, IReadOnlyList<StudioNativeRectangle> LineBounds);

internal sealed record StudioNativeRectangle(double X, double Y, double Width, double Height) {
    internal static StudioNativeRectangle From(StudioRectangle rectangle) => new(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
}
