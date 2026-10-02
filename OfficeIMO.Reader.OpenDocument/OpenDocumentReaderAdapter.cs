using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Native OpenDocument ingestion adapter for <see cref="OfficeDocumentReader"/>.</summary>
internal static partial class OpenDocumentReaderAdapter {
    private const int MaximumTableColumns = 256;

    /// <summary>Reads an ODT, ODS, or ODP file and emits format-aligned chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? options = null, ReaderOpenDocumentOptions? openDocumentOptions = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException($"OpenDocument file '{path}' doesn't exist.", path);
        ReaderOptions effective = options ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effective.MaxInputBytes);
        ReaderOpenDocumentOptions formatOptions = (openDocumentOptions ?? new ReaderOpenDocumentOptions()).Clone();
        OdfDocument document = OdfDocument.Load(path, CreateOpenOptions(effective, formatOptions));
        foreach (ReaderChunk chunk in ReadDocument(document, Path.GetFullPath(path), effective, formatOptions, cancellationToken)) yield return chunk;
    }

    /// <summary>Reads an ODT, ODS, or ODP stream without closing the caller's stream.</summary>
    public static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? options = null, ReaderOpenDocumentOptions? openDocumentOptions = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OpenDocument stream must be readable.", nameof(stream));
        ReaderOptions effective = options ?? new ReaderOptions();
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, effective.MaxInputBytes, cancellationToken, out bool ownsStream);
        try {
            ReaderOpenDocumentOptions formatOptions = (openDocumentOptions ?? new ReaderOpenDocumentOptions()).Clone();
            OdfDocument document = OdfDocument.Load(parseStream, CreateOpenOptions(effective, formatOptions));
            string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "document.odf" : sourceName!.Trim();
            foreach (ReaderChunk chunk in ReadDocument(document, logicalName, effective, formatOptions, cancellationToken)) yield return chunk;
        } finally {
            if (ownsStream) parseStream.Dispose();
        }
    }

    private static IEnumerable<ReaderChunk> ReadDocument(OdfDocument document, string sourceName, ReaderOptions options, ReaderOpenDocumentOptions formatOptions,
        CancellationToken cancellationToken) {
        foreach (ReaderChunk source in ReadUnboundedDocument(document, sourceName, options, formatOptions, new ProjectionBudget(formatOptions.MaxExtractedCharacters), cancellationToken)) {
            int limit = options.MaxChars > 0 ? options.MaxChars : 8000;
            if (Math.Max(source.Text?.Length ?? 0, source.Markdown?.Length ?? 0) <= limit) {
                yield return source;
                continue;
            }
            ReaderChunkHierarchyResult split = ReaderHierarchicalChunker.Chunk(new[] { source }, new ReaderHierarchicalChunkingOptions {
                MaxTokens = limit, OverlapTokens = 0, MaxInputChunks = 1,
                IncludeContextInText = false, TokenCounter = CharacterCounter.Instance
            }, cancellationToken);
            if (split.Diagnostics.Any(diagnostic => diagnostic.Code == "hierarchical-output-chunk-limit")) {
                throw new InvalidDataException("OpenDocument character chunking exceeded the output chunk limit.");
            }
            foreach (ReaderChunk chunk in split.Chunks) {
                chunk.TokenEstimate = ReaderHeuristicTokenCounter.Instance.CountTokens(chunk.Text ?? string.Empty);
                yield return chunk;
            }
        }
    }

    private sealed class ProjectionBudget {
        private long _remaining;
        internal ProjectionBudget(long maximum) {
            if (maximum < 1) throw new ArgumentOutOfRangeException(nameof(ReaderOpenDocumentOptions.MaxExtractedCharacters));
            _remaining = maximum;
        }
        internal string Read(string text) {
            if (text.Length > _remaining) throw new InvalidDataException("OpenDocument extraction exceeds MaxExtractedCharacters.");
            _remaining -= text.Length;
            return text;
        }
    }

    private sealed class CharacterCounter : IReaderTokenCounter {
        internal static readonly CharacterCounter Instance = new CharacterCounter();
        public string Id => "officeimo.reader.characters";
        public int CountTokens(string text) => text.Length;
    }

    private static IEnumerable<ReaderChunk> ReadUnboundedDocument(OdfDocument document, string sourceName, ReaderOptions options, ReaderOpenDocumentOptions formatOptions, ProjectionBudget budget,
        CancellationToken cancellationToken) {
        if (document is OdtDocument text) {
            foreach (ReaderChunk chunk in ReadTextDocument(text, sourceName, options, budget, cancellationToken)) yield return chunk;
        } else if (document is OdsDocument spreadsheet) {
            foreach (ReaderChunk chunk in ReadSpreadsheet(spreadsheet, sourceName, options, formatOptions, budget, cancellationToken)) yield return chunk;
        } else if (document is OdpPresentation presentation) {
            foreach (ReaderChunk chunk in ReadPresentation(presentation, sourceName, options, formatOptions, budget, cancellationToken)) yield return chunk;
        }
    }

    private static OdfLoadOptions CreateOpenOptions(ReaderOptions options, ReaderOpenDocumentOptions formatOptions) {
        var result = new OdfLoadOptions { Password = options.OpenPassword };
        if (options.MaxInputBytes.HasValue) result.MaxPackageBytes = options.MaxInputBytes.Value;
        if (formatOptions.MaxXmlCharacters.HasValue) result.MaxXmlCharacters = formatOptions.MaxXmlCharacters.Value;
        return result;
    }

    private static string BuildId(string sourceName, string kind, int index, int part = 0) {
        string safe = new string(Path.GetFileNameWithoutExtension(sourceName).ToLowerInvariant()
            .Select(character => char.IsLetterOrDigit(character) ? character : '-')
            .ToArray()).Trim('-');
        if (safe.Length == 0) safe = "document";
        return string.Concat("odf-", safe, "-", kind, "-", index.ToString("D4", CultureInfo.InvariantCulture), "-", part.ToString("D3", CultureInfo.InvariantCulture));
    }
}
