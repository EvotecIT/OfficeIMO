using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

/// <summary>
/// Projects an already loaded Visio document into the shared OfficeIMO reader model.
/// </summary>
public static class VisioDocumentReaderExtensions {
    /// <summary>
    /// Converts an already loaded Visio document into the shared structured read result without serializing and reopening it.
    /// </summary>
    public static OfficeDocumentReadResult ToOfficeDocumentReadResult(
        this VisioDocument document,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderVisioOptions? visioOptions = null,
        CancellationToken cancellationToken = default) {
        ReaderOptions effectiveReaderOptions =
            DocumentReaderEngine.NormalizeOptions(readerOptions);
        return VisioReaderAdapter.ReadDocument(
            document,
            sourceName,
            effectiveReaderOptions,
            visioOptions,
            cancellationToken);
    }
}
