namespace OfficeIMO.Reader.Latex;

/// <summary>LaTeX ingestion entry points.</summary>
public static class DocumentReaderLatexExtensions {
    /// <summary>Reads a `.tex` file.</summary>
    public static IEnumerable<ReaderChunk> ReadLatexFile(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderLatexOptions? latexOptions = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("LaTeX file does not exist.", path);
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        ReaderLatexOptions adapter = ReaderLatexOptionsCloner.Clone(latexOptions);
        LatexParseResult result = LatexDocument.Parse(File.ReadAllText(path), adapter.ParseOptions);
        return ReadResult(result, path, reader, adapter, cancellationToken);
    }

    /// <summary>Reads a caller-owned LaTeX stream.</summary>
    public static IEnumerable<ReaderChunk> ReadLatex(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderLatexOptions? latexOptions = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("LaTeX stream must be readable.", nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderLatexOptions adapter = ReaderLatexOptionsCloner.Clone(latexOptions);
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, reader.MaxInputBytes, cancellationToken, out bool ownsStream);
        try {
            using var textReader = new StreamReader(parseStream, Encoding.UTF8, true, 4096, leaveOpen: true);
            LatexParseResult result = LatexDocument.Parse(textReader.ReadToEnd(), adapter.ParseOptions);
            string name = string.IsNullOrWhiteSpace(sourceName) ? "document.tex" : sourceName!.Trim();
            return ReadResult(result, name, reader, adapter, cancellationToken).ToArray();
        } finally {
            if (ownsStream) parseStream.Dispose();
        }
    }

    /// <summary>Adapts an already parsed document.</summary>
    public static IEnumerable<ReaderChunk> ReadLatexDocument(
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
