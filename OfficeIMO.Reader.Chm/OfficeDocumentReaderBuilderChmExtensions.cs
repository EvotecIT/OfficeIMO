namespace OfficeIMO.Reader.Chm;

/// <summary>Registers in-process compiled help ingestion in an isolated reader.</summary>
public static class OfficeDocumentReaderBuilderChmExtensions {
    /// <summary>Stable CHM handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.chm";
    /// <summary>Adds CHM topics, tables, links and resources with archive-based source hashes and per-topic citations.</summary>
    public static OfficeDocumentReaderBuilder AddChmHandler(this OfficeDocumentReaderBuilder builder,
        ReaderChmOptions? options = null, bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderChmOptions registered = options?.Clone() ?? new ReaderChmOptions();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Id = HandlerId, Origin = ReaderHandlerOrigin.OfficeIMO, Kind = ReaderInputKind.Chm,
            DisplayName = "Compiled HTML Help Reader", Description = "Bounded CHM archive and HTML topic ingestion.", Extensions = new[] { ".chm" },
            DefaultMaxInputBytes = registered.ReadOptions.MaxInputBytes, MaxInputBytesCeiling = registered.ReadOptions.MaxInputBytes,
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged, WarningBehavior = ReaderWarningBehavior.Mixed, DeterministicOutput = true,
            ReadPath = (path, reader, token) => ChmReaderAdapter.ReadDocument(path, reader, registered.Clone(), token).Chunks,
            ReadStream = (stream, name, reader, token) => ChmReaderAdapter.ReadDocument(stream, name, reader, registered.Clone(), token).Chunks,
            ReadDocumentPath = (path, reader, token) => ChmReaderAdapter.ReadDocument(path, reader, registered.Clone(), token),
            ReadDocumentStream = (stream, name, reader, token) => ChmReaderAdapter.ReadDocument(stream, name, reader, registered.Clone(), token)
        }, replaceExisting);
    }
}
