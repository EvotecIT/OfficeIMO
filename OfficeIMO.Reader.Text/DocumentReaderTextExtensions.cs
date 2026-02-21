using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Xml;

namespace OfficeIMO.Reader.Text;

/// <summary>
/// Structured text orchestration helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderTextExtensions {
    /// <summary>
    /// Reads structured text content with CSV/JSON/XML-aware chunking.
    /// Other inputs fallback to <see cref="DocumentReader.Read(string, ReaderOptions?, CancellationToken)"/>.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadStructuredText(string path, ReaderOptions? readerOptions = null, StructuredTextReadOptions? structuredOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var options = Normalize(structuredOptions);
        var ext = GetNormalizedExtension(path);

        if (ext == ".csv" || ext == ".tsv") {
            foreach (var chunk in DocumentReaderCsvExtensions.ReadCsv(path, readerOptions, ToCsvOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".json") {
            foreach (var chunk in DocumentReaderJsonExtensions.ReadJson(path, readerOptions, ToJsonOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".xml") {
            foreach (var chunk in DocumentReaderXmlExtensions.ReadXml(path, readerOptions, ToXmlOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        foreach (var chunk in DocumentReader.Read(path, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads structured text content from a stream with CSV/JSON/XML-aware chunking.
    /// Other inputs fallback to <see cref="DocumentReader.Read(Stream, string?, ReaderOptions?, CancellationToken)"/>.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadStructuredText(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, StructuredTextReadOptions? structuredOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var options = Normalize(structuredOptions);
        var ext = GetNormalizedExtension(sourceName);

        if (ext == ".csv" || ext == ".tsv") {
            foreach (var chunk in DocumentReaderCsvExtensions.ReadCsv(stream, sourceName, readerOptions, ToCsvOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".json") {
            foreach (var chunk in DocumentReaderJsonExtensions.ReadJson(stream, sourceName, readerOptions, ToJsonOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".xml") {
            foreach (var chunk in DocumentReaderXmlExtensions.ReadXml(stream, sourceName, readerOptions, ToXmlOptions(options), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        foreach (var chunk in DocumentReader.Read(stream, sourceName, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static CsvReadOptions ToCsvOptions(StructuredTextReadOptions options) {
        return new CsvReadOptions {
            ChunkRows = options.CsvChunkRows,
            HeadersInFirstRow = options.CsvHeadersInFirstRow,
            IncludeMarkdown = options.IncludeCsvMarkdown
        };
    }

    private static JsonReadOptions ToJsonOptions(StructuredTextReadOptions options) {
        return new JsonReadOptions {
            ChunkRows = options.JsonChunkRows,
            MaxDepth = options.JsonMaxDepth,
            IncludeMarkdown = options.IncludeJsonMarkdown
        };
    }

    private static XmlReadOptions ToXmlOptions(StructuredTextReadOptions options) {
        return new XmlReadOptions {
            ChunkRows = options.XmlChunkRows,
            IncludeMarkdown = options.IncludeXmlMarkdown
        };
    }

    private static StructuredTextReadOptions Normalize(StructuredTextReadOptions? options) {
        var source = options ?? new StructuredTextReadOptions();

        var normalized = new StructuredTextReadOptions {
            CsvChunkRows = source.CsvChunkRows,
            CsvHeadersInFirstRow = source.CsvHeadersInFirstRow,
            IncludeCsvMarkdown = source.IncludeCsvMarkdown,
            JsonChunkRows = source.JsonChunkRows,
            JsonMaxDepth = source.JsonMaxDepth,
            IncludeJsonMarkdown = source.IncludeJsonMarkdown,
            XmlChunkRows = source.XmlChunkRows,
            IncludeXmlMarkdown = source.IncludeXmlMarkdown
        };

        if (normalized.CsvChunkRows < 1) normalized.CsvChunkRows = 1;
        if (normalized.JsonChunkRows < 1) normalized.JsonChunkRows = 1;
        if (normalized.XmlChunkRows < 1) normalized.XmlChunkRows = 1;
        if (normalized.JsonMaxDepth < 1) normalized.JsonMaxDepth = 1;

        return normalized;
    }

    private static string GetNormalizedExtension(string? sourceName) {
        var ext = Path.GetExtension(sourceName ?? string.Empty);
        return ext.ToLowerInvariant();
    }
}
