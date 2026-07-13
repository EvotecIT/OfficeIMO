namespace OfficeIMO.Reader.AsciiDoc;

/// <summary>AsciiDoc ingestion entry points for <see cref="OfficeDocumentReader"/>.</summary>
internal static class AsciiDocReaderAdapter {
    /// <summary>Reads an AsciiDoc file into source-aware Reader chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderAsciiDocOptions? asciiDocOptions = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("AsciiDoc path cannot be empty.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("AsciiDoc file does not exist.", path);

        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        string source = File.ReadAllText(path);
        AsciiDocParseResult result = AsciiDocDocument.Parse(source, ReaderAsciiDocOptionsCloner.Clone(asciiDocOptions).ParseOptions);
        return ReadAsciiDocResult(result, path, reader, asciiDocOptions, cancellationToken);
    }

    /// <summary>Reads an AsciiDoc stream without closing the caller-owned stream.</summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderAsciiDocOptions? asciiDocOptions = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("AsciiDoc stream must be readable.", nameof(stream));

        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderAsciiDocOptions adapter = ReaderAsciiDocOptionsCloner.Clone(asciiDocOptions);
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, reader.MaxInputBytes, cancellationToken, out bool ownsStream);
        try {
            using var textReader = new StreamReader(parseStream, Encoding.UTF8, true, 4096, leaveOpen: true);
            string source = textReader.ReadToEnd();
            cancellationToken.ThrowIfCancellationRequested();
            AsciiDocParseResult result = AsciiDocDocument.Parse(source, adapter.ParseOptions);
            string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "document.adoc" : sourceName!.Trim();
            return ReadAsciiDocResult(result, logicalName, reader, adapter, cancellationToken).ToArray();
        } finally {
            if (ownsStream) parseStream.Dispose();
        }
    }

    /// <summary>Adapts an already parsed native document to Reader chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(
        AsciiDocDocument document,
        string sourceName = "document.adoc",
        ReaderOptions? readerOptions = null,
        ReaderAsciiDocOptions? asciiDocOptions = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        var result = new AsciiDocParseResult(document, document.Diagnostics);
        return ReadAsciiDocResult(result, sourceName, readerOptions ?? new ReaderOptions(), asciiDocOptions, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadAsciiDocResult(
        AsciiDocParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderAsciiDocOptions? asciiDocOptions,
        CancellationToken cancellationToken) {
        ReaderAsciiDocOptions options = ReaderAsciiDocOptionsCloner.Clone(asciiDocOptions);
        return options.ChunkByBlock
            ? AsciiDocReaderChunkBuilder.BuildBlockChunks(result, sourceName, readerOptions, options, cancellationToken)
            : AsciiDocReaderChunkBuilder.BuildDocumentChunks(result, sourceName, readerOptions, options, cancellationToken);
    }
}
