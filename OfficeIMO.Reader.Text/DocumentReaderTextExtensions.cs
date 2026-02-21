using OfficeIMO.CSV;

namespace OfficeIMO.Reader.Text;

/// <summary>
/// Structured text ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderTextExtensions {
    /// <summary>
    /// Reads structured text content with CSV-aware chunking.
    /// Non-CSV inputs fallback to <see cref="DocumentReader.Read(string, ReaderOptions?, CancellationToken)"/>.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadStructuredText(string path, ReaderOptions? readerOptions = null, StructuredTextReadOptions? structuredOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is ".csv" or ".tsv") {
            foreach (var chunk in ReadCsv(path, ext, structuredOptions ?? new StructuredTextReadOptions(), cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        foreach (var chunk in DocumentReader.Read(path, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsv(string path, string extension, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        if (options.CsvChunkRows < 1) options.CsvChunkRows = 1;

        var delimiter = extension == ".tsv" ? '\t' : ',';
        var csv = CsvDocument.Load(path, new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = options.CsvHeadersInFirstRow,
            Mode = CsvLoadMode.Stream
        });

        var headers = csv.Header.Count > 0
            ? csv.Header.ToArray()
            : new[] { "Column1" };

        var rows = new List<IReadOnlyList<string>>(capacity: options.CsvChunkRows);
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

            if (rows.Count >= options.CsvChunkRows) {
                yield return BuildCsvChunk(path, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeCsvMarkdown);
                rows = new List<IReadOnlyList<string>>(capacity: options.CsvChunkRows);
                chunkIndex++;
            }
        }

        if (rows.Count > 0) {
            yield return BuildCsvChunk(path, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeCsvMarkdown);
        }
    }

    private static ReaderChunk BuildCsvChunk(string path, IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows, int chunkIndex, int sourceRowStart, bool includeMarkdown) {
        var table = new ReaderTable {
            Title = Path.GetFileName(path),
            Columns = headers.ToArray(),
            Rows = rows.Select(static r => (IReadOnlyList<string>)r.ToArray()).ToArray(),
            TotalRowCount = rows.Count,
            Truncated = false
        };

        var plain = BuildPlain(headers, rows);
        var markdown = includeMarkdown ? BuildMarkdown(headers, rows) : null;

        return new ReaderChunk {
            Id = string.Concat("csv-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
            Kind = ReaderInputKind.Text,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = sourceRowStart
            },
            Text = plain,
            Markdown = markdown,
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
        sb.Append(string.Join(" | ", headers));
        sb.AppendLine(" |");

        sb.Append("| ");
        sb.Append(string.Join(" | ", headers.Select(static _ => "---")));
        sb.AppendLine(" |");

        foreach (var row in rows) {
            sb.Append("| ");
            sb.Append(string.Join(" | ", row));
            sb.AppendLine(" |");
        }

        return sb.ToString().TrimEnd();
    }
}
