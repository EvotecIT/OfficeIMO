using System.Globalization;
using OfficeIMO.Pdf;

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
