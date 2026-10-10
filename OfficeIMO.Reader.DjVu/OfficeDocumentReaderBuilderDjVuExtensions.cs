namespace OfficeIMO.Reader.DjVu;

/// <summary>Adds native managed DjVu ingestion to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderDjVuExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.djvu";
    /// <summary>Registers bounded path and stream ingestion for .djvu and .djv.</summary>
    public static OfficeDocumentReaderBuilder AddDjVuHandler(this OfficeDocumentReaderBuilder builder,
        ReaderDjVuOptions? options = null, bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        var snapshot = (options ?? new ReaderDjVuOptions()).Clone();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO, Id = HandlerId, DisplayName = "DjVu Reader Adapter",
            Description = "Managed DjVu stored text, native geometry and optional complete page images.", Kind = ReaderInputKind.DjVu,
            Extensions = new[] { ".djvu", ".djv" }, DefaultMaxInputBytes = snapshot.ReadOptions.MaxSourceBytes,
            MaxInputBytesCeiling = snapshot.ReadOptions.MaxSourceBytes,
            ReadPath = (path, reader, token) => DjVuReaderAdapter.Read(path, reader, snapshot.Clone(), token).Chunks,
            ReadStream = (stream, name, reader, token) => DjVuReaderAdapter.Read(stream, name, reader, snapshot.Clone(), token).Chunks,
            ReadDocumentPath = (path, reader, token) => DjVuReaderAdapter.Read(path, reader, snapshot.Clone(), token),
            ReadDocumentStream = (stream, name, reader, token) => DjVuReaderAdapter.Read(stream, name, reader, snapshot.Clone(), token),
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            WarningBehavior = ReaderWarningBehavior.Mixed, DeterministicOutput = true,
            FormatQualifications = new[] { Profile(".djvu"), Profile(".djv") }
        }, replaceExisting);
    }
    private static ReaderFormatQualification Profile(string extension) => new ReaderFormatQualification(extension, "djvu", ReaderFormatSupport.ReadConvert,
        "DjVu v3 stored-text and scanned-page projection", new[] { "Page order and component identity", "Stored UTF-8 text and word geometry", "Explicit missing/corrupt text status", "Optional display-oriented PNG page assets" },
        new[] { "No DjVu authoring or active annotations", "Indirect components require a trusted explicit resolver", "Raster assets and OCR are opt-in" },
        new[] { "OfficeIMO.DjVu/SUPPORT.md", "OfficeIMO.DjVu.Tests" });
}
