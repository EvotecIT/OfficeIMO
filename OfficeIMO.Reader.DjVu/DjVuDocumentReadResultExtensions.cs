using OfficeIMO.DjVu;

namespace OfficeIMO.Reader.DjVu;

/// <summary>Projects an already owned DjVu document for the shared Reader and OCR pipeline.</summary>
public static class DjVuDocumentReadResultExtensions {
    /// <summary>Returns stored text, display geometry and optional raster assets without reloading the source or running OCR.</summary>
    public static OfficeDocumentReadResult ToReadResult(this DjVuDocument document, ReaderDjVuOptions? options = null,
        ReaderOptions? readerOptions = null, string? sourceName = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var reader = readerOptions ?? ReaderOptions.CreateSafeIngestion();
        string name = sourceName ?? "document.djvu";
        return DjVuReaderAdapter.Project(document, new OfficeDocumentSource { Path = name,
            SourceId = DocumentReaderEngine.BuildSourceId(name), SourceHash = reader.ComputeHashes ? document.SourceSha256 : null,
            LengthBytes = document.SourceLengthBytes },
            reader, (options ?? new ReaderDjVuOptions()).Clone(), cancellationToken);
    }
}
