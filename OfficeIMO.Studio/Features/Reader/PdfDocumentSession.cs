using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Owns one immutable OfficeIMO PDF snapshot and exposes reader-oriented operations to the desktop host.
/// </summary>
internal sealed class PdfDocumentSession {
    private readonly PdfDocument _document;

    private PdfDocumentSession(string path, long fileSize, PdfDocument document, PdfDocumentInfo documentInfo) {
        Path = path;
        FileSize = fileSize;
        _document = document;
        DocumentInfo = documentInfo;
    }

    internal string Path { get; }

    internal string FileName => System.IO.Path.GetFileName(Path);

    internal long FileSize { get; }

    internal PdfDocumentInfo DocumentInfo { get; }

    internal IReadOnlyList<PdfPageInfo> Pages => DocumentInfo.Pages;

    internal async Task<IReadOnlyList<PdfSearchHit>> SearchAsync(
        string query,
        CancellationToken cancellationToken,
        IProgress<double>? progress = null) {
        if (string.IsNullOrWhiteSpace(query)) return Array.Empty<PdfSearchHit>();
        string needle = query.Trim();
        return await Task.Run<IReadOnlyList<PdfSearchHit>>(() => {
            var matches = new List<PdfSearchHit>();
            int pageCount = Pages.Count;
            for (int index = 0; index < pageCount; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<string> pageText = _document.Read.TextByPage(
                    (index + 1).ToString(CultureInfo.InvariantCulture));
                cancellationToken.ThrowIfCancellationRequested();
                string text = pageText.Count == 0 ? string.Empty : pageText[0];
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
        return new PdfDocumentSession(
            workspace.Path,
            workspace.FileSize,
            workspace.CreateDocumentSnapshot(),
            workspace.DocumentInfo);
    }

    internal async Task<PdfPageScene> LoadPageSceneAsync(
        int pageNumber,
        CancellationToken cancellationToken) {
        if (pageNumber <= 0 || pageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        return await Task.Run(() => {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeIMO.Drawing.OfficeDrawing drawing = _document.Read.Drawing(pageNumber);
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageInteractionMap interactions = _document.Read.Interactions(pageNumber);
            IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics =
                _document.Read.RenderCapabilityDiagnostics(pageNumber);
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
            .OpenAsync(fullPath, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        PdfDocumentInfo documentInfo = await Task
            .Run(() => document.Read.DocumentInfo(), cancellationToken)
            .ConfigureAwait(false);

        cancellationToken.ThrowIfCancellationRequested();
        return new PdfDocumentSession(fullPath, file.Length, document, documentInfo);
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

        IReadOnlyList<PdfPageRenderResult> results = await Task.Run(
            () => _document.Read.RenderPages(
                pageNumber.ToString(CultureInfo.InvariantCulture),
                options,
                cancellationToken: cancellationToken),
            cancellationToken).ConfigureAwait(false);

        PdfPageRenderResult result = results.Count == 1
            ? results[0]
            : throw new InvalidOperationException("The PDF renderer did not return the requested page.");

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
