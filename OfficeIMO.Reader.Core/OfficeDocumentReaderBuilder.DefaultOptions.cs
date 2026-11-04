using System;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReaderBuilder {
    private ReaderOptions? _defaultReadOptions;

    /// <summary>Captures the default read and folder policies for readers built from this builder.</summary>
    /// <remarks>Both registration and Build capture independent snapshots. Null per-read options use this profile; an explicit per-read object replaces it.</remarks>
    public OfficeDocumentReaderBuilder WithDefaultOptions(ReaderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        _defaultReadOptions = DocumentReaderEngine.NormalizeOptions(options, inheritProjection: false);
        return this;
    }
}
