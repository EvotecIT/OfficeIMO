using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Owns one immutable OfficeIMO PDF snapshot and exposes reader-oriented operations to the desktop host.
/// </summary>
internal sealed class PdfDocumentSession {
    private readonly PdfDocument _document;
    private readonly object _displayLock = new();
    private PdfPageDisplayRenderer? _display;
    private readonly bool _canEditAnnotations;

    private PdfDocumentSession(
        string path,
        string fileName,
        long fileSize,
        PdfDocument document,
        PdfDocumentViewInfo viewInfo,
        bool canEditAnnotations = false) {
        Path = path;
        FileName = fileName;
        FileSize = fileSize;
        _document = document;
        ViewInfo = viewInfo;
        _canEditAnnotations = canEditAnnotations;
    }

    internal string Path { get; }

    internal string FileName { get; }

    internal long FileSize { get; }

    internal PdfDocumentViewInfo ViewInfo { get; }

    internal PdfDocumentInfo? DocumentInfo => ViewInfo.LogicalContent;

    internal bool CanSearch => ViewInfo.CanExtractText;

    // Search finds a scan's OCR layer, so selection and copying reach it too.
    private static readonly PdfPageInteractionOptions SelectionOptions = new() { IncludeTextRenderingMode3 = true };

    internal IReadOnlyList<PdfPageInfo> Pages => ViewInfo.Pages;

    internal PdfVisualComparisonReport CompareTo(PdfDocumentSession actual, PdfVisualComparisonOptions options, CancellationToken token) =>
        _document.Proof.CompareVisual(actual._document, token, options: options);

    /// <summary>
    /// A page's text elements in extraction order, each with its geometry on the page as displayed: from its top-left
    /// corner, after the page box, rotation and user unit. Hosts select and copy text from them.
    /// </summary>
    internal IReadOnlyList<PdfPageInteractionRegion> GetTextRegions(int pageNumber, CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (!CanSearch) throw new NotSupportedException("Copying text is restricted by this document's permissions.");
        cancellationToken.ThrowIfCancellationRequested();
        return _document.Render.Interactions(pageNumber, SelectionOptions).TextRegions;
    }

    internal async Task<IReadOnlyList<PdfSearchHit>> SearchAsync(
        string query,
        CancellationToken cancellationToken,
        IProgress<double>? progress = null) {
        if (string.IsNullOrWhiteSpace(query)) return Array.Empty<PdfSearchHit>();
        if (!CanSearch) throw new InvalidOperationException("Text search is restricted by this document's permissions.");
        string needle = query.Trim();
        return await Task.Run<IReadOnlyList<PdfSearchHit>>(() => {
            var matches = new List<PdfSearchHit>();
            int pageCount = Pages.Count;
            cancellationToken.ThrowIfCancellationRequested();
            var pageMatches = _document.Text.Find(needle, new PdfTextSearchOptions { IncludeTextRenderingMode3 = true })
                .GroupBy(static match => match.PageNumber)
                .ToDictionary(static group => group.Key, static group => group.ToArray());
            cancellationToken.ThrowIfCancellationRequested();
            for (int index = 0; index < pageCount; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (pageMatches.TryGetValue(index + 1, out var occurrences)) {
                    foreach (var occurrence in occurrences) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var bounds = occurrence.VisualBounds;
                        matches.Add(new PdfSearchHit(index + 1, occurrence.Text) {
                            Bounds = new StudioRectangle(bounds.Left, bounds.Top, bounds.Width, bounds.Height),
                            LineBounds = occurrence.VisualLineBounds
                                .Select(static line => new StudioRectangle(line.Left, line.Top, line.Width, line.Height))
                                .ToArray(),
                            OccurrenceNumber = matches.Count + 1
                        });
                    }
                }
                progress?.Report((index + 1D) / pageCount);
            }
            cancellationToken.ThrowIfCancellationRequested();
            return matches.AsReadOnly();
        }, cancellationToken).ConfigureAwait(false);
    }

    internal static PdfDocumentSession FromWorkspace(PdfWorkspace workspace) {
        ArgumentNullException.ThrowIfNull(workspace);
        PdfDocument document = workspace.CreateDocumentSnapshot();
        return new PdfDocumentSession(
            workspace.Path,
            workspace.FileName,
            workspace.FileSize,
            document,
            workspace.ViewInfo,
            workspace.CanEditAnnotations);
    }

    internal async Task<PdfPageScene> LoadPageSceneAsync(
        int pageNumber,
        CancellationToken cancellationToken,
        Func<OfficeIMO.Drawing.OfficeDrawing, IReadOnlyList<string>>? analyzePresentation = null) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        return await Task.Run(() => {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ViewInfo.CanExtractContent) {
                (double width, double height) = GetDisplaySize(pageNumber);
                var display = new OfficeIMO.Drawing.OfficeDrawing(width, height);
                PdfPageInteractionMap? annotationInteractions = _canEditAnnotations
                    ? _document.Annotations.GetEditingInteractions(pageNumber) : null;
                return new PdfPageScene(pageNumber, display, annotationInteractions, [], RequiresRasterFallback: true);
            }
            OfficeIMO.Drawing.OfficeDrawing drawing = _document.Render.Drawing(pageNumber);
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageInteractionMap interactions = _document.Render.Interactions(pageNumber, SelectionOptions);
            IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics =
                _document.Render.CapabilityDiagnostics(pageNumber);
            IReadOnlyList<string> adapterDiagnostics =
                analyzePresentation?.Invoke(drawing) ?? [];
            cancellationToken.ThrowIfCancellationRequested();

            return new PdfPageScene(
                pageNumber,
                drawing,
                interactions,
                diagnostics.Select(static diagnostic => diagnostic.Message).Concat(adapterDiagnostics).ToArray(),
                adapterDiagnostics.Count > 0) {
                Capabilities = diagnostics
            };
        }, cancellationToken).ConfigureAwait(false);
    }

    internal static async Task<PdfDocumentSession> OpenAsync(string path, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("A PDF path is required.", nameof(path));
        }

        string fullPath = System.IO.Path.GetFullPath(path);
        if (!File.Exists(fullPath)) {
            throw new FileNotFoundException("The selected PDF no longer exists.", fullPath);
        }

        if (!string.Equals(System.IO.Path.GetExtension(fullPath), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException("OfficeIMO Studio currently opens PDF documents.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        var file = new FileInfo(fullPath);
        PdfDocument document = await PdfDocument
            .LoadAsync(fullPath, StudioPdfSecurityPolicy.CreateLoadOptions(), cancellationToken)
            .ConfigureAwait(false);

        PdfDocumentViewInfo documentInfo = await Task
            .Run(() => document.InspectForViewing(cancellationToken: cancellationToken), cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfDocumentSession(fullPath, file.Name, file.Length, document, documentInfo);
    }

    /// <summary>The page as displayed, in points: its effective box at the page's user unit, rotated.</summary>
    internal (double Width, double Height) GetDisplaySize(int pageNumber) {
        PdfPageInfo info = Pages[pageNumber - 1];
        (double width, double height) = GetPhysicalSize(info);
        return Math.Abs(info.RotationDegrees) % 180 == 90 ? (height, width) : (width, height);
    }

    /// <summary>The page's size in points (1/72 inch) before rotation: its effective box at the page's user unit.</summary>
    internal static (double Width, double Height) GetPhysicalSize(PdfPageInfo info) =>
        ((info.Geometry.EffectiveBox?.Width ?? info.Width) * (info.UserUnit ?? 1D),
         (info.Geometry.EffectiveBox?.Height ?? info.Height) * (info.UserUnit ?? 1D));

    /// <summary>
    /// Renders a page as PNG. The page image stays within <see cref="StudioPdfSecurityPolicy.MaximumRasterPixels"/>;
    /// a larger scale is refused before any page content is read.
    /// </summary>
    internal async Task<PdfRenderedPage> RenderPageAsync(
        int pageNumber,
        double scale,
        CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }
        (double displayWidth, double displayHeight) = GetDisplaySize(pageNumber);
        double pixels = Math.Ceiling(displayWidth * scale) * Math.Ceiling(displayHeight * scale);
        if (!(pixels <= StudioPdfSecurityPolicy.MaximumRasterPixels)) {
            throw new ArgumentOutOfRangeException(nameof(scale), scale,
                $"At this scale the page needs {pixels:N0} pixels; pages render at up to {StudioPdfSecurityPolicy.MaximumRasterPixels:N0}.");
        }

        var options = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = scale,
            MaxPages = 1,
            // One page: its failure, such as a read or pixel limit, reaches the host as the typed exception.
            ContinueOnError = false,
            MaxPixelsPerPage = StudioPdfSecurityPolicy.MaximumRenderWorkingPixels,
            RenderTimeout = StudioPdfSecurityPolicy.RenderTimeout,
            MaxTotalOutputBytes = StudioPdfSecurityPolicy.MaximumRasterOutputBytes,
            MaxOutputBytesPerPage = StudioPdfSecurityPolicy.MaximumRasterOutputBytes
        };

        PdfPageRenderResult result;
        if (!ViewInfo.CanExtractContent) {
            result = await Task.Run(() => _document.Render.DisplayPage(pageNumber,
                new PdfPageDisplayOptions {
                    Scale = scale,
                    MaximumPixels = StudioPdfSecurityPolicy.MaximumRenderWorkingPixels,
                    MaximumOutputBytes = options.MaxOutputBytesPerPage,
                    Timeout = StudioPdfSecurityPolicy.RenderTimeout
                }, cancellationToken),
                cancellationToken).ConfigureAwait(false);
        } else {
            IReadOnlyList<PdfPageRenderResult> results = await Task.Run(
                () => _document.Render.Pages(pageNumber.ToString(CultureInfo.InvariantCulture), options, cancellationToken),
                cancellationToken).ConfigureAwait(false);
            result = results.Count == 1 ? results[0]
                : throw new InvalidOperationException("The PDF renderer did not return the requested page.");
        }

        byte[]? bytes = result.Bytes;
        if (!result.Succeeded || bytes is null) {
            string detail = result.Diagnostics.Count == 0
                ? "The managed renderer could not render this page."
                : string.Join(Environment.NewLine, result.Diagnostics);
            throw new InvalidOperationException(detail);
        }

        return new PdfRenderedPage(
            result.PageNumber,
            scale,
            bytes,
            result.Width,
            result.Height,
            result.Elapsed,
            result.Diagnostics);
    }

    /// <summary>
    /// Renders a rectangle of a page as opaque RGBA pixels for display. The page image at <paramref name="scale"/> is
    /// the page's display size times the scale, each side rounded up; the rectangle may be all of it or a tile, and
    /// tiles join without seams. A rectangle holds at most <see cref="StudioPdfSecurityPolicy.MaximumRasterPixels"/>.
    /// Recently rendered pages and their decoded images are kept for further tiles and zoom levels of this revision.
    /// </summary>
    internal async Task<PdfRenderedPixels> RenderRegionAsync(
        int pageNumber, double scale, int x, int y, int width, int height, CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (!(scale > 0D) || double.IsInfinity(scale))
            throw new ArgumentOutOfRangeException(nameof(scale), scale, "The scale must be a finite positive number.");
        if (width <= 0 || height <= 0 || (long)width * height > StudioPdfSecurityPolicy.MaximumRasterPixels) {
            throw new ArgumentOutOfRangeException(nameof(width),
                $"A rectangle holds 1 to {StudioPdfSecurityPolicy.MaximumRasterPixels:N0} pixels; render a larger area as tiles.");
        }
        var limits = new PdfPageRegionLimits {
            MaximumRegionPixels = StudioPdfSecurityPolicy.MaximumRasterPixels,
            MaximumWorkingPixels = StudioPdfSecurityPolicy.MaximumRenderWorkingPixels,
            Timeout = StudioPdfSecurityPolicy.RenderTimeout
        };
        OfficeIMO.Drawing.OfficeRasterImage image = await Task.Run(() => {
            lock (_displayLock) {
                _display ??= new PdfPageDisplayRenderer(_document, StudioPdfSecurityPolicy.DisplayCachedPages,
                    StudioPdfSecurityPolicy.DisplayImageCacheBytes);
                return _display.RenderRegion(pageNumber, scale, x, y, width, height, limits, cancellationToken);
            }
        }, cancellationToken).ConfigureAwait(false);
        return new PdfRenderedPixels(pageNumber, scale, x, y, image.Width, image.Height, image.PixelBuffer);
    }
}

/// <summary>Opaque RGBA pixels of a page rectangle, rows top to bottom, four bytes per pixel.</summary>
internal sealed record PdfRenderedPixels(int PageNumber, double Scale, int X, int Y, int Width, int Height, byte[] Pixels);
