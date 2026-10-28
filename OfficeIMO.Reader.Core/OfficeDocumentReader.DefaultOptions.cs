namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    private readonly ReaderOptions _defaultReadOptions;

    /// <summary>Returns an independent copy of this reader's default read and folder policies.</summary>
    /// <remarks>Changing the returned object does not reconfigure the reader. An explicit per-read options object replaces the default read profile.</remarks>
    public ReaderOptions DefaultOptions => DocumentReaderEngine.NormalizeOptions(_defaultReadOptions, inheritProjection: false);

    private ReaderOptions NormalizeReadOptions(ReaderOptions? options) =>
        DocumentReaderEngine.NormalizeOptions(options ?? _defaultReadOptions);

}
