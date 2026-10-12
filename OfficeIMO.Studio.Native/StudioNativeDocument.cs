using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
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
                StudioPdfSecurityPolicy.MaximumRasterPixels, _workspace.CanMutatePages, _workspace.CanExtractPages, _workspace.CanEditAnnotations,
                !view.Security.HasEncryption || view.Security.HasOwnerAuthorization || view.Security.AllowsPrinting == true);
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

    // The markup tools a host may add. Text, images, stamps, links and redaction have their own journeys.
    private static readonly HashSet<PdfEditorTool> AnnotationTools = [
        PdfEditorTool.Ink, PdfEditorTool.Note, PdfEditorTool.FreeText, PdfEditorTool.Highlight, PdfEditorTool.Underline,
        PdfEditorTool.StrikeOut, PdfEditorTool.Squiggly, PdfEditorTool.Rectangle, PdfEditorTool.Ellipse, PdfEditorTool.Line,
        PdfEditorTool.Polygon
    ];

    /// <summary>
    /// Adds a markup annotation as one undoable change. Geometry is in points of the page as displayed, from its
    /// top-left corner; the workspace maps it to the page's own space through its box, rotation and user unit.
    /// </summary>
    internal Task AddAnnotationAsync(StudioNativeAnnotationRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        if (!Enum.TryParse(request.Tool, ignoreCase: true, out PdfEditorTool tool) || !AnnotationTools.Contains(tool))
            throw new ArgumentException($"'{request.Tool}' is not a markup tool.", nameof(request));
        var strokes = (request.Strokes ?? []).Select(stroke => (IReadOnlyList<PdfEditorVisualPoint>)stroke
            .Select(point => new PdfEditorVisualPoint(Finite(point.X), Finite(point.Y))).ToArray()).Where(stroke => stroke.Count > 0).ToArray();
        var quads = (request.Quads ?? []).Select(quad => new PdfEditorVisualBounds(Finite(quad.X), Finite(quad.Y),
            Finite(quad.X + quad.Width), Finite(quad.Y + quad.Height))).ToArray();
        double? width = request.StrokeWidth;
        if (width is { } w && (!double.IsFinite(w) || w <= 0D || w > 144D)) throw new ArgumentOutOfRangeException(nameof(request), "A stroke is wider than 0 and at most 144 points.");
        if (request.Opacity is { } o && (!double.IsFinite(o) || o <= 0D || o > 1D)) throw new ArgumentOutOfRangeException(nameof(request), "Opacity is above 0 and at most 1.");
        (double left, double top, double right, double bottom) = request.Bounds is { } bounds
            ? (Finite(bounds.X), Finite(bounds.Y), Finite(bounds.X + bounds.Width), Finite(bounds.Y + bounds.Height))
            : BoundsOf(strokes.SelectMany(static stroke => stroke), (width ?? 1D) / 2D + 1D);
        if (right <= left || bottom <= top) throw new ArgumentException("The markup needs an area or a drawn path.", nameof(request));
        StudioNativeColor color = request.Color ?? new StudioNativeColor(1D, 0.85D, 0D);
        if (new[] { color.Red, color.Green, color.Blue }.Any(static part => !double.IsFinite(part) || part < 0D || part > 1D))
            throw new ArgumentOutOfRangeException(nameof(request), "Colour components are from 0 through 1.");
        var gesture = new PdfEditorGesture(request.Page, left, top, right, bottom, strokes.FirstOrDefault() ?? [],
            strokes.Length > 0 ? strokes : null, quads.Length > 0 ? quads : null);
        var properties = new PdfEditorProperties(request.Text ?? string.Empty, request.Author ?? string.Empty,
            new PdfColor(color.Red, color.Green, color.Blue), string.Empty, string.Empty, request.FontSize ?? 12D,
            StrokeWidth: width, Opacity: request.Opacity);
        return RunAsync(() => {
            if (request.Page < 1 || request.Page > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(request), "The page does not exist.");
            if (!_workspace.CanEditAnnotations) throw new NotSupportedException("This document's annotations can't be changed.");
            return _workspace.ApplyEditorGestureAsync(tool, gesture, properties, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>The markup annotations on a page, with rectangles in points of the page as displayed.</summary>
    internal async Task<StudioNativeAnnotationList> GetAnnotationsAsync(int pageNumber, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            var metadata = _workspace.AnnotationMetadata.Where(static annotation => annotation.ObjectNumber.HasValue)
                .GroupBy(static annotation => annotation.ObjectNumber!.Value).ToDictionary(static group => group.Key, static group => group.First());
            PdfPageInteractionMap map = _workspace.CreateDocumentSnapshot().Annotations.GetEditingInteractions(pageNumber);
            var annotations = new List<StudioNativeAnnotation>();
            foreach (PdfPageInteractionRegion region in map.Regions) {
                cancellationToken.ThrowIfCancellationRequested();
                if (region.Kind != PdfInteractionKind.Annotation || region.ObjectNumber is not { } number) continue;
                metadata.TryGetValue(number, out PdfAnnotation? annotation);
                StudioNativeColor? color = annotation?.Color is { Count: 3 } rgb ? new StudioNativeColor(rgb[0], rgb[1], rgb[2]) : null;
                annotations.Add(new StudioNativeAnnotation(number, pageNumber, region.Subtype ?? annotation?.Subtype ?? string.Empty,
                    new StudioNativeRectangle(region.Quad.Left, region.Quad.Top, region.Quad.Width, region.Quad.Height),
                    annotation?.Contents, annotation?.Title, color));
            }
            return new StudioNativeAnnotationList(StudioNativeAbi.Version, _workspace.Revision, pageNumber, _workspace.CanEditAnnotations, annotations);
        } finally { _gate.Release(); }
    }

    /// <summary>Changes an annotation's text, author or colour as one undoable change; what is not given is kept.</summary>
    internal Task UpdateAnnotationAsync(StudioNativeAnnotationUpdate update, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(update);
        return RunAsync(() => {
            PdfAnnotation current = _workspace.AnnotationMetadata.FirstOrDefault(annotation => annotation.ObjectNumber == update.ObjectNumber)
                ?? throw new ArgumentException("The annotation does not exist.", nameof(update));
            PdfColor color = update.Color is { } given ? new PdfColor(given.Red, given.Green, given.Blue)
                : current.Color is { Count: 3 } rgb ? new PdfColor(rgb[0], rgb[1], rgb[2]) : new PdfColor(1D, 0.85D, 0D);
            return _workspace.UpdateAnnotationAsync(update.ObjectNumber, update.Contents ?? current.Contents ?? string.Empty,
                update.Author ?? current.Title ?? string.Empty, color, cancellationToken);
        }, cancellationToken);
    }

    /// <summary>Removes an annotation as one undoable change.</summary>
    internal Task RemoveAnnotationAsync(int objectNumber, CancellationToken cancellationToken) =>
        RunAsync(() => {
            if (!_workspace.AnnotationMetadata.Any(annotation => annotation.ObjectNumber == objectNumber))
                throw new ArgumentException("The annotation does not exist.", nameof(objectNumber));
            return _workspace.RemoveAnnotationAsync(objectNumber, cancellationToken);
        }, cancellationToken);

    private static double Finite(double value) =>
        double.IsFinite(value) ? value : throw new ArgumentOutOfRangeException(nameof(value), "Coordinates must be finite numbers.");

    private static (double Left, double Top, double Right, double Bottom) BoundsOf(IEnumerable<PdfEditorVisualPoint> points, double margin) {
        double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
        foreach (PdfEditorVisualPoint point in points) {
            left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
        }
        return left > right ? (0D, 0D, 0D, 0D) : (left - margin, top - margin, right + margin, bottom + margin);
    }

    /// <summary>
    /// A page's text elements for selection and copying: each element's text, its box, as x, y, width and height in
    /// points of the page as displayed, four numbers per element in the same order, and how it follows the element before
    /// it as a <see cref="PdfTextBreak"/>, so words and lines are found as the reader finds them.
    /// </summary>
    internal async Task<StudioNativePageText> GetPageTextAsync(int pageNumber, CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (pageNumber < 1 || pageNumber > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber), "The page does not exist.");
            PdfDocumentSession session = CurrentSession();
            IReadOnlyList<PdfPageInteractionRegion> regions = await Task.Run(() => session.GetTextRegions(pageNumber, cancellationToken), cancellationToken)
                .ConfigureAwait(false);
            var texts = new string[regions.Count];
            var boxes = new double[regions.Count * 4];
            var breaks = new int[regions.Count];
            for (int index = 0; index < regions.Count; index++) {
                breaks[index] = (int)(index == 0 ? PdfTextBreak.Line : PdfPageInteractionMap.GetTextBreak(regions[index - 1], regions[index]));
                PdfSelectionQuad quad = regions[index].Quad;
                texts[index] = regions[index].Text ?? string.Empty;
                boxes[index * 4] = Math.Round(quad.Left, 2);
                boxes[index * 4 + 1] = Math.Round(quad.Top, 2);
                boxes[index * 4 + 2] = Math.Round(quad.Width, 2);
                boxes[index * 4 + 3] = Math.Round(quad.Height, 2);
            }
            return new StudioNativePageText(StudioNativeAbi.Version, _workspace.Revision, pageNumber, texts, boxes, breaks);
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
    bool CanExtractPages, bool CanEditAnnotations, bool CanPrint);

internal sealed record StudioNativePageInfo(int Number, double Width, double Height, int Rotation);

internal sealed record StudioNativeSearchResult(int AbiVersion, string Query, long Revision, IReadOnlyList<StudioNativeSearchHit> Hits);

internal sealed record StudioNativeSearchHit(int Page, int Occurrence, string Text, StudioNativeRectangle Bounds, IReadOnlyList<StudioNativeRectangle> LineBounds);

internal sealed record StudioNativeRectangle(double X, double Y, double Width, double Height) {
    internal static StudioNativeRectangle From(StudioRectangle rectangle) => new(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
}

internal sealed record StudioNativePoint(double X, double Y);

internal sealed record StudioNativeColor(double Red, double Green, double Blue);

internal sealed record StudioNativeAnnotationRequest(string Tool, int Page, StudioNativeRectangle? Bounds = null,
    IReadOnlyList<IReadOnlyList<StudioNativePoint>>? Strokes = null, IReadOnlyList<StudioNativeRectangle>? Quads = null,
    string? Text = null, string? Author = null, StudioNativeColor? Color = null, double? StrokeWidth = null, double? Opacity = null,
    double? FontSize = null);

internal sealed record StudioNativeAnnotation(int ObjectNumber, int Page, string Subtype, StudioNativeRectangle Bounds, string? Contents,
    string? Author, StudioNativeColor? Color);

internal sealed record StudioNativeAnnotationList(int AbiVersion, long Revision, int Page, bool CanEdit, IReadOnlyList<StudioNativeAnnotation> Annotations);

internal sealed record StudioNativeAnnotationUpdate(int ObjectNumber, string? Contents = null, string? Author = null, StudioNativeColor? Color = null);

internal sealed record StudioNativePageText(int AbiVersion, long Revision, int Page, IReadOnlyList<string> Texts, IReadOnlyList<double> Boxes,
    IReadOnlyList<int> Breaks);
