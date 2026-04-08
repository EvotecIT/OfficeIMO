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

        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);
        var options = Normalize(csvOptions);
        var extension = GetNormalizedExtension(path);
        var delimiter = extension == ".tsv" ? '\t' : ',';

        var csv = CsvDocument.Load(path, new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = options.HeadersInFirstRow,
            Mode = CsvLoadMode.Stream
        });

        foreach (var chunk in ReadCsvDocument(csv, source, options, effectiveReaderOptions.ComputeHashes, cancellationToken)) {
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
            var source = BuildSourceMetadataFromStream(parseStream, sourcePath, effectiveReaderOptions.ComputeHashes);
            var csv = CsvDocument.Load(parseStream, new CsvLoadOptions {
                Delimiter = delimiter,
                HasHeaderRow = options.HeadersInFirstRow,
                Mode = CsvLoadMode.Stream
            }, leaveOpen: true);

            foreach (var chunk in ReadCsvDocument(csv, source, options, effectiveReaderOptions.ComputeHashes, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsvDocument(CsvDocument csv, SourceMetadata source, CsvReadOptions options, bool computeHashes, CancellationToken cancellationToken) {
        var sourcePath = source.Path;
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
                yield return EnrichChunk(BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeMarkdown), source, computeHashes);
                rows = new List<IReadOnlyList<string>>(capacity: options.ChunkRows);
                chunkIndex++;
            }
        }

        if (rows.Count > 0) {
            yield return EnrichChunk(BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeMarkdown), source, computeHashes);
            yield break;
        }

        if (rowIndex == 0) {
            if (csv.Header.Count > 0) {
                yield return EnrichChunk(
                    BuildCsvChunk(
                        sourcePath,
                        headers,
                        Array.Empty<IReadOnlyList<string>>(),
                        chunkIndex,
                        sourceRowStart: 0,
                        includeMarkdown: options.IncludeMarkdown),
                    source,
                    computeHashes);
            } else {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "csv-warning-0000", "CSV content produced no rows."), source, computeHashes);
            }
        }
    }

    private static ReaderChunk BuildCsvChunk(
        string path,
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<string>> rows,
        int chunkIndex,
        int sourceRowStart,
        bool includeMarkdown) {
        var effectiveHeaders = ExpandHeaders(headers, rows);
        var normalizedRows = NormalizeRows(rows, effectiveHeaders.Length);
        var table = new ReaderTable {
            Title = Path.GetFileName(path),
            Columns = effectiveHeaders,
            Rows = normalizedRows,
            TotalRowCount = normalizedRows.Count,
            Truncated = false
        };

        return new ReaderChunk {
            Id = "csv-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
            Kind = ReaderInputKind.Csv,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = sourceRowStart
            },
            Text = BuildPlain(effectiveHeaders, normalizedRows),
            Markdown = includeMarkdown ? BuildMarkdown(effectiveHeaders, normalizedRows) : null,
            Tables = new[] { table }
        };
    }

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Csv,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = 0
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static string[] ExpandHeaders(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
        int maxColumns = headers.Count;
        foreach (var row in rows) {
            if (row.Count > maxColumns) {
                maxColumns = row.Count;
            }
        }

        var expanded = new string[maxColumns];
        for (int i = 0; i < maxColumns; i++) {
            if (i < headers.Count && !string.IsNullOrWhiteSpace(headers[i])) {
                expanded[i] = headers[i].Trim();
            } else {
                expanded[i] = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
            }
        }

        EnsureUniqueHeaders(expanded);
        return expanded;
    }

    private static IReadOnlyList<IReadOnlyList<string>> NormalizeRows(IReadOnlyList<IReadOnlyList<string>> rows, int columnCount) {
        var normalized = new IReadOnlyList<string>[rows.Count];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            if (row.Count == columnCount) {
                normalized[rowIndex] = row.ToArray();
                continue;
            }

            var values = new string[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                values[columnIndex] = columnIndex < row.Count ? row[columnIndex] ?? string.Empty : string.Empty;
            }

            normalized[rowIndex] = values;
        }

        return normalized;
    }

    private static void EnsureUniqueHeaders(string[] headers) {
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++) {
            var original = string.IsNullOrWhiteSpace(headers[i])
                ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture)
                : headers[i];

            if (!seen.TryAdd(original, 1)) {
                var suffix = seen[original] + 1;
                string candidate;
                do {
                    candidate = original + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                } while (seen.ContainsKey(candidate));

                seen[original] = suffix - 1;
                headers[i] = candidate;
                seen[candidate] = 1;
                continue;
            }

            headers[i] = original;
        }
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
        var ext = Path.GetExtension((sourceName ?? string.Empty).Trim());
        return ext.ToLowerInvariant();
    }

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName.Trim();
        }

        return defaultName;
    }

    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceMetadata source, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (source == null) throw new ArgumentNullException(nameof(source));

        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        if (!chunk.TokenEstimate.HasValue) {
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        }
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }

        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        var safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        var data = string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.HeadingPath ?? string.Empty,
            chunk.Location.HeadingSlug ?? string.Empty,
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Location.Sheet ?? string.Empty,
            chunk.Location.A1Range ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.Slide?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.StartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedStartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedEndLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty);

        return ComputeSha256Hex(data);
    }

    private static SourceMetadata BuildSourceMetadataFromPath(string path, bool computeHash) {
        var normalizedPath = NormalizePathForId(path);
        var sourceId = BuildSourceId(normalizedPath);

        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fileInfo = new FileInfo(path);
            if (fileInfo.Exists) {
                lastWriteUtc = fileInfo.LastWriteTimeUtc;
                lengthBytes = fileInfo.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeFileSha256(path);
        }

        return new SourceMetadata {
            Path = path,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static SourceMetadata BuildSourceMetadataFromStream(Stream stream, string sourceName, bool computeHash) {
        var logicalName = string.IsNullOrWhiteSpace(sourceName) ? "document.csv" : sourceName.Trim();
        var sourceId = BuildSourceId(logicalName);

        long? lengthBytes = null;
        try {
            if (stream.CanSeek) {
                lengthBytes = stream.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeStreamSha256(stream);
        }

        return new SourceMetadata {
            Path = logicalName,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = null,
            LengthBytes = lengthBytes
        };
    }

    private static string? TryComputeFileSha256(string path) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs);
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream) {
        if (stream == null || !stream.CanSeek) return null;

        long position;
        try {
            position = stream.Position;
        } catch {
            return null;
        }

        try {
            stream.Position = 0;
            var hash = ComputeSha256Hex(stream);
            stream.Position = position;
            return hash;
        } catch {
            try {
                stream.Position = position;
            } catch {
                // ignore
            }

            return null;
        }
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
        var hash = sha.ComputeHash(bytes);
        return ConvertToHexLower(hash);
    }

    private static string ComputeSha256Hex(Stream stream) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return ConvertToHexLower(hash);
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static string BuildSourceId(string sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar == '\\') {
            normalized = normalized.ToLowerInvariant();
        }

        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        string fullPath;
        try {
            fullPath = Path.GetFullPath(path);
        } catch {
            fullPath = path;
        }

        return fullPath.Replace('\\', '/');
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
