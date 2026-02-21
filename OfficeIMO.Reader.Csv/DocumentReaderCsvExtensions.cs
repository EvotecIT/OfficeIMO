using OfficeIMO.CSV;

namespace OfficeIMO.Reader.Csv;

/// <summary>
/// CSV/TSV ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderCsvExtensions {
    /// <summary>
    /// Reads CSV/TSV content from a path with table-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadCsv(string path, ReaderOptions? readerOptions = null, CsvReadOptions? csvOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        var options = Normalize(csvOptions);
        var extension = GetNormalizedExtension(path);
        var delimiter = extension == ".tsv" ? '\t' : ',';

        var csv = CsvDocument.Load(path, new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = options.HeadersInFirstRow,
            Mode = CsvLoadMode.Stream
        });

        foreach (var chunk in ReadCsvDocument(csv, path, options, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads CSV/TSV content from a stream with table-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadCsv(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, CsvReadOptions? csvOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var options = Normalize(csvOptions);
        var extension = GetNormalizedExtension(sourceName);

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);

        try {
            var delimiter = extension == ".tsv" ? '\t' : ',';
            var sourcePath = BuildLogicalSourcePath(sourceName, extension == ".tsv" ? "document.tsv" : "document.csv");
            var csv = CsvDocument.Load(parseStream, new CsvLoadOptions {
                Delimiter = delimiter,
                HasHeaderRow = options.HeadersInFirstRow,
                Mode = CsvLoadMode.Stream
            }, leaveOpen: true);

            foreach (var chunk in ReadCsvDocument(csv, sourcePath, options, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsvDocument(CsvDocument csv, string sourcePath, CsvReadOptions options, CancellationToken cancellationToken) {
        var headers = csv.Header.Count > 0
            ? csv.Header.ToArray()
            : new[] { "Column1" };

        var rows = new List<IReadOnlyList<string>>(capacity: options.ChunkRows);
        int chunkIndex = 0;
        int rowIndex = 0;

        foreach (var row in csv.AsEnumerable()) {
            cancellationToken.ThrowIfCancellationRequested();

            var values = new string[Math.Max(headers.Length, row.FieldCount)];
            for (int i = 0; i < values.Length; i++) {
                values[i] = i < row.FieldCount
                    ? Convert.ToString(row[i], CultureInfo.InvariantCulture) ?? string.Empty
                    : string.Empty;
            }

            rows.Add(values);
            rowIndex++;

            if (rows.Count >= options.ChunkRows) {
                yield return BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeMarkdown);
                rows = new List<IReadOnlyList<string>>(capacity: options.ChunkRows);
                chunkIndex++;
            }
        }

        if (rows.Count > 0) {
            yield return BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeMarkdown);
        }
    }

    private static ReaderChunk BuildCsvChunk(
        string path,
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<string>> rows,
        int chunkIndex,
        int sourceRowStart,
        bool includeMarkdown) {
        var table = new ReaderTable {
            Title = Path.GetFileName(path),
            Columns = headers.ToArray(),
            Rows = rows.Select(static r => (IReadOnlyList<string>)r.ToArray()).ToArray(),
            TotalRowCount = rows.Count,
            Truncated = false
        };

        return new ReaderChunk {
            Id = "csv-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
            Kind = ReaderInputKind.Text,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = sourceRowStart
            },
            Text = BuildPlain(headers, rows),
            Markdown = includeMarkdown ? BuildMarkdown(headers, rows) : null,
            Tables = new[] { table }
        };
    }

    private static string BuildPlain(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(" | ", headers));
        foreach (var row in rows) {
            sb.AppendLine(string.Join(" | ", row));
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildMarkdown(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
        var sb = new StringBuilder();

        sb.Append("| ");
        sb.Append(string.Join(" | ", headers.Select(EscapeMarkdownCell)));
        sb.AppendLine(" |");

        sb.Append("| ");
        sb.Append(string.Join(" | ", headers.Select(static _ => "---")));
        sb.AppendLine(" |");

        foreach (var row in rows) {
            sb.Append("| ");
            sb.Append(string.Join(" | ", row.Select(EscapeMarkdownCell)));
            sb.AppendLine(" |");
        }

        return sb.ToString().TrimEnd();
    }

    private static string EscapeMarkdownCell(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Replace("\\", "\\\\").Replace("|", "\\|");
    }

    private static CsvReadOptions Normalize(CsvReadOptions? options) {
        var source = options ?? new CsvReadOptions();

        var normalized = new CsvReadOptions {
            ChunkRows = source.ChunkRows,
            HeadersInFirstRow = source.HeadersInFirstRow,
            IncludeMarkdown = source.IncludeMarkdown
        };

        if (normalized.ChunkRows < 1) {
            normalized.ChunkRows = 1;
        }

        return normalized;
    }

    private static string GetNormalizedExtension(string? sourceName) {
        var ext = Path.GetExtension(sourceName ?? string.Empty);
        return ext.ToLowerInvariant();
    }

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName!;
        }

        return defaultName;
    }
}
