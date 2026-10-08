namespace OfficeIMO.Reader.Xps;

/// <summary>Adds native XPS/OpenXPS support to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderXpsExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.xps";
    /// <summary>Registers bounded path and stream ingestion for .xps and .oxps.</summary>
    public static OfficeDocumentReaderBuilder AddXpsHandler(this OfficeDocumentReaderBuilder builder,
        ReaderXpsOptions? options = null, bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        var snapshot = (options ?? new ReaderXpsOptions()).Clone();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO, Id = HandlerId, DisplayName = "XPS Reader Adapter",
            Description = "Native XPS/OpenXPS Unicode and logical structure through OfficeIMO.Xps.", Kind = ReaderInputKind.Xps,
            Extensions = new[] { ".xps", ".oxps" }, DefaultMaxInputBytes = snapshot.ReadOptions.MaximumInputBytes,
            MaxInputBytesCeiling = snapshot.ReadOptions.MaximumInputBytes,
            ReadPath = (path, reader, token) => XpsReaderAdapter.Read(path, reader, snapshot.Clone(), token).Chunks,
            ReadStream = (stream, name, reader, token) => XpsReaderAdapter.Read(stream, name, reader, snapshot.Clone(), token).Chunks,
            ReadDocumentPath = (path, reader, token) => XpsReaderAdapter.Read(path, reader, snapshot.Clone(), token),
            ReadDocumentStream = (stream, name, reader, token) => XpsReaderAdapter.Read(stream, name, reader, snapshot.Clone(), token),
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            WarningBehavior = ReaderWarningBehavior.Mixed, DeterministicOutput = true
        }, replaceExisting);
    }
}
