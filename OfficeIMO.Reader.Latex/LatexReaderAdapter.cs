namespace OfficeIMO.Reader.Latex;

/// <summary>LaTeX ingestion entry points.</summary>
internal static class LatexReaderAdapter {
    /// <summary>Reads a `.tex` file.</summary>
    public static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderLatexOptions? latexOptions = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("LaTeX file does not exist.", path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Read(stream, path, readerOptions, latexOptions, cancellationToken);
    }

    /// <summary>Reads a caller-owned LaTeX stream.</summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderLatexOptions? latexOptions = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderLatexOptions adapter = ReaderLatexOptionsCloner.Clone(latexOptions);
        long? nativeLimit = adapter.ParseOptions.MaximumInputBytes;
        adapter.ParseOptions.MaximumInputBytes = reader.MaxInputBytes.HasValue
            ? nativeLimit.HasValue ? Math.Min(nativeLimit.Value, reader.MaxInputBytes.Value) : reader.MaxInputBytes
            : nativeLimit;
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, adapter.ParseOptions.MaximumInputBytes, cancellationToken, out bool ownsStream);
        try {
            LatexParseResult result = LatexDocument.LoadResult(parseStream, adapter.ParseOptions, null, cancellationToken);
            string name = string.IsNullOrWhiteSpace(sourceName) ? "document.tex" : sourceName!.Trim();
            return ReadResult(result, name, reader, adapter, cancellationToken).ToArray();
        } finally {
            if (ownsStream) parseStream.Dispose();
        }
    }

    /// <summary>Adapts an already parsed document.</summary>
    public static IEnumerable<ReaderChunk> Read(
        LatexDocument document,
        string sourceName = "document.tex",
        ReaderOptions? readerOptions = null,
        ReaderLatexOptions? latexOptions = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadResult(new LatexParseResult(document, document.Diagnostics), sourceName,
            readerOptions ?? new ReaderOptions(), ReaderLatexOptionsCloner.Clone(latexOptions), cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadResult(
        LatexParseResult result,
        string sourceName,
        ReaderOptions reader,
        ReaderLatexOptions options,
        CancellationToken cancellationToken) =>
        options.ChunkByBlock
            ? LatexReaderChunkBuilder.BuildBlocks(result, sourceName, reader, options, cancellationToken)
            : LatexReaderChunkBuilder.BuildDocument(result, sourceName, reader, options, cancellationToken);
}
