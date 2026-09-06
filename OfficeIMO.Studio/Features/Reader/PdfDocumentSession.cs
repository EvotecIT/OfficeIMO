using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Owns one immutable OfficeIMO PDF snapshot and exposes reader-oriented operations to the desktop host.
/// </summary>
internal sealed class PdfDocumentSession {
    private readonly PdfDocument _document;
    private readonly PdfDocumentReadResult? _semanticDocument;

    private PdfDocumentSession(
        string path,
        string fileName,
        long fileSize,
        PdfDocument document,
        PdfDocumentViewInfo viewInfo,
        PdfDocumentReadResult? semanticDocument) {
        Path = path;
        FileName = fileName;
        FileSize = fileSize;
        _document = document;
        ViewInfo = viewInfo;
        _semanticDocument = semanticDocument;
    }

    internal string Path { get; }

    internal string FileName { get; }

    internal long FileSize { get; }

    internal PdfDocumentViewInfo ViewInfo { get; }

    internal PdfDocumentInfo? DocumentInfo => ViewInfo.LogicalContent;

    internal bool CanSearch => ViewInfo.CanExtractText;

    internal IReadOnlyList<PdfPageInfo> Pages => ViewInfo.Pages;

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
            // The canonical text search authenticates and parses once for the entire snapshot.
            // It permits accessibility extraction without exposing a logical document or interaction map.
            var restrictedMatches = _semanticDocument is null
                ? _document.Text.Find(needle, new PdfTextSearchOptions { IncludeTextRenderingMode3 = true })
                    .GroupBy(static match => match.PageNumber)
                    .ToDictionary(static group => group.Key, static group => group.First())
                : null;
            cancellationToken.ThrowIfCancellationRequested();
            for (int index = 0; index < pageCount; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_semanticDocument is null) {
                    if (restrictedMatches!.TryGetValue(index + 1, out var restrictedMatch))
                        matches.Add(new PdfSearchHit(index + 1, restrictedMatch.Text));
                    progress?.Report((index + 1D) / pageCount);
                    continue;
                }
                PdfLogicalPage? page = _semanticDocument.Pages.FirstOrDefault(
                    candidate => candidate.PageNumber == index + 1);
                cancellationToken.ThrowIfCancellationRequested();
                string text = page is null
                    ? string.Empty
                    : string.Join(Environment.NewLine, page.TextBlocks.Select(static block => block.Text));
                int match = text.IndexOf(needle, StringComparison.OrdinalIgnoreCase);
                if (match >= 0) {
                    int start = Math.Max(0, match - 32);
                    int length = Math.Min(text.Length - start, needle.Length + 64);
                    string snippet = string.Join(" ", text.Substring(start, length)
                        .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
                    matches.Add(new PdfSearchHit(index + 1, snippet));
                }
                progress?.Report((index + 1D) / pageCount);
            }
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
            workspace.ViewInfo.CanExtractContent ? document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }) : null);
    }

    internal async Task<PdfPageScene> LoadPageSceneAsync(
        int pageNumber,
        CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        return await Task.Run(() => {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ViewInfo.CanExtractContent) {
                PdfPageInfo info = Pages[pageNumber - 1];
                bool rotated = Math.Abs(info.RotationDegrees) % 180 == 90;
                var display = new OfficeIMO.Drawing.OfficeDrawing(rotated ? info.Height : info.Width, rotated ? info.Width : info.Height);
                return new PdfPageScene(pageNumber, display, null, [], RequiresRasterFallback: true);
            }
            OfficeIMO.Drawing.OfficeDrawing drawing = _document.Render.Drawing(pageNumber);
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageInteractionMap interactions = _document.Render.Interactions(pageNumber);
            IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics =
                _document.Render.CapabilityDiagnostics(pageNumber);
            IReadOnlyList<string> adapterDiagnostics =
                OfficeDrawingAvaloniaRenderer.AnalyzeRasterFallback(drawing);
            cancellationToken.ThrowIfCancellationRequested();

            return new PdfPageScene(
                pageNumber,
                drawing,
                interactions,
                diagnostics.Select(static diagnostic => diagnostic.Message).Concat(adapterDiagnostics).ToArray(),
                adapterDiagnostics.Count > 0);
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
            .LoadAsync(fullPath, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        PdfDocumentViewInfo documentInfo = await Task
            .Run(() => document.InspectForViewing(cancellationToken: cancellationToken), cancellationToken)
            .ConfigureAwait(false);
        PdfDocumentReadResult? semanticDocument = documentInfo.CanExtractContent ? await Task
            .Run(
                () => document.Read(
                    new PdfReadOptions { Profile = PdfReadProfile.Fast },
                    cancellationToken),
                cancellationToken)
            .ConfigureAwait(false) : null;

        cancellationToken.ThrowIfCancellationRequested();
        return new PdfDocumentSession(fullPath, file.Name, file.Length, document, documentInfo, semanticDocument);
    }

    internal async Task<PdfRenderedPage> RenderPageAsync(
        int pageNumber,
        double scale,
        CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        var options = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = scale,
            MaxPages = 1,
            ContinueOnError = true,
            MaxTotalOutputBytes = 64L * 1024L * 1024L,
            MaxOutputBytesPerPage = 64L * 1024L * 1024L
        };

        PdfPageRenderResult result;
        if (!ViewInfo.CanExtractContent) {
            result = await Task.Run(() => _document.Render.DisplayPage(pageNumber,
                new PdfPageDisplayOptions { Scale = scale, MaximumOutputBytes = options.MaxOutputBytesPerPage }, cancellationToken),
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
}
