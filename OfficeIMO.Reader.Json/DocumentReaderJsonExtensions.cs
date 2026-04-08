namespace OfficeIMO.Reader.Json;

/// <summary>
/// JSON ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderJsonExtensions {
    /// <summary>
    /// Reads JSON content from a path with AST-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadJson(string path, ReaderOptions? readerOptions = null, JsonReadOptions? jsonOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadJson(fs, source, effectiveReaderOptions, jsonOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads JSON content from a stream with AST-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadJson(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, JsonReadOptions? jsonOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var options = Normalize(jsonOptions);
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.json");
        return ReadJson(stream, BuildSourceMetadataFromLogicalStream(stream, sourcePath, effectiveReaderOptions, cancellationToken), effectiveReaderOptions, options, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadJson(Stream stream, SourceMetadata source, ReaderOptions effectiveReaderOptions, JsonReadOptions? jsonOptions, CancellationToken cancellationToken) {
        var options = Normalize(jsonOptions);
        var sourcePath = source.Path;

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);
        UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);

        JsonDocument? doc = null;
        string? parseError = null;
        try {
            try {
                doc = JsonDocument.Parse(parseStream, new JsonDocumentOptions {
                    AllowTrailingCommas = true,
                    CommentHandling = JsonCommentHandling.Skip,
                    MaxDepth = options.MaxDepth
                });
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                parseError = "JSON parse error: " + ex.GetType().Name + ".";
            }

            if (parseError != null) {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "json-warning-0000", parseError), source, effectiveReaderOptions.ComputeHashes);
                yield break;
            }

            using (doc) {
                var rows = new List<StructuredRow>(capacity: 1024);
                TraverseJson(doc!.RootElement, "$", depth: 0, options.MaxDepth, rows, cancellationToken);

                foreach (var chunk in BuildStructuredChunks(source, rows, options, effectiveReaderOptions.ComputeHashes, cancellationToken)) {
                    yield return chunk;
                }
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> BuildStructuredChunks(
        SourceMetadata source,
        IReadOnlyList<StructuredRow> rows,
        JsonReadOptions options,
        bool computeHashes,
        CancellationToken cancellationToken) {
        if (rows.Count == 0) {
            yield break;
        }

        var path = source.Path;
        int index = 0;
        int chunkIndex = 0;
        while (index < rows.Count) {
            cancellationToken.ThrowIfCancellationRequested();

            int take = Math.Min(options.ChunkRows, rows.Count - index);
            var slice = new List<StructuredRow>(take);
            for (int i = 0; i < take; i++) {
                slice.Add(rows[index + i]);
            }

            var tableRows = slice
                .Select(static r => (IReadOnlyList<string>)new[] { r.Path, r.Type, r.Value })
                .ToArray();

            var table = new ReaderTable {
                Title = Path.GetFileName(path),
                Columns = new[] { "Path", "Type", "Value" },
                Rows = tableRows,
                TotalRowCount = slice.Count,
                Truncated = false
            };

            yield return EnrichChunk(new ReaderChunk {
                Id = "json-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Json,
                Location = new ReaderLocation {
                    Path = path,
                    BlockIndex = chunkIndex,
                    SourceBlockIndex = index
                },
                Text = BuildPlain(slice),
                Markdown = options.IncludeMarkdown ? BuildMarkdown(slice) : null,
                Tables = new[] { table }
            }, source, computeHashes);

            index += take;
            chunkIndex++;
        }
    }

    private static void TraverseJson(
        JsonElement element,
        string path,
        int depth,
        int maxDepth,
        List<StructuredRow> rows,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        if (depth > maxDepth) {
            rows.Add(new StructuredRow(path, "depth-limit", "(max depth reached)"));
            return;
        }

        switch (element.ValueKind) {
            case JsonValueKind.Object:
                if (!element.EnumerateObject().Any()) {
                    rows.Add(new StructuredRow(path, "object", "{}"));
                }

                foreach (var property in element.EnumerateObject()) {
                    var childPath = AppendPropertyPath(path, property.Name);
                    TraverseJson(property.Value, childPath, depth + 1, maxDepth, rows, cancellationToken);
                }

                break;
            case JsonValueKind.Array:
                if (element.GetArrayLength() == 0) {
                    rows.Add(new StructuredRow(path, "array", "[]"));
                }

                int index = 0;
                foreach (var child in element.EnumerateArray()) {
                    var childPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                    TraverseJson(child, childPath, depth + 1, maxDepth, rows, cancellationToken);
                    index++;
                }

                break;
            case JsonValueKind.String:
                rows.Add(new StructuredRow(path, "string", element.GetString() ?? string.Empty));
                break;
            case JsonValueKind.Number:
                rows.Add(new StructuredRow(path, "number", element.GetRawText()));
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                rows.Add(new StructuredRow(path, "boolean", element.GetRawText()));
                break;
            case JsonValueKind.Null:
                rows.Add(new StructuredRow(path, "null", "null"));
                break;
            default:
                rows.Add(new StructuredRow(path, "value", element.GetRawText()));
                break;
        }
    }

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Json,
            Location = new ReaderLocation { Path = path },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static string BuildPlain(IReadOnlyList<StructuredRow> rows) {
        var sb = new StringBuilder();
        sb.AppendLine("Path | Type | Value");
        foreach (var row in rows) {
            sb.Append(row.Path);
            sb.Append(" | ");
            sb.Append(row.Type);
            sb.Append(" | ");
            sb.AppendLine(row.Value);
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildMarkdown(IReadOnlyList<StructuredRow> rows) {
        var sb = new StringBuilder();
        sb.AppendLine("| Path | Type | Value |");
        sb.AppendLine("| --- | --- | --- |");
        foreach (var row in rows) {
            sb.Append("| ");
            sb.Append(EscapeMarkdownCell(row.Path));
            sb.Append(" | ");
            sb.Append(EscapeMarkdownCell(row.Type));
            sb.Append(" | ");
            sb.Append(EscapeMarkdownCell(row.Value));
            sb.AppendLine(" |");
        }

        return sb.ToString().TrimEnd();
    }

    private static string EscapeMarkdownCell(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Replace("\\", "\\\\").Replace("|", "\\|");
    }

    private static JsonReadOptions Normalize(JsonReadOptions? options) {
        var source = options ?? new JsonReadOptions();

        var normalized = new JsonReadOptions {
            ChunkRows = source.ChunkRows,
            MaxDepth = source.MaxDepth,
            IncludeMarkdown = source.IncludeMarkdown
        };

        if (normalized.ChunkRows < 1) normalized.ChunkRows = 1;
        if (normalized.MaxDepth < 1) normalized.MaxDepth = 1;

        return normalized;
    }

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                return trimmedSourceName;
            }
        }

        return defaultName;
    }

    private static string AppendPropertyPath(string path, string propertyName) {
        if (IsSimpleJsonPathIdentifier(propertyName)) {
            return path == "$"
                ? "$." + propertyName
                : path + "." + propertyName;
        }

        return path + "[\"" + EscapeJsonPathString(propertyName) + "\"]";
    }

    private static bool IsSimpleJsonPathIdentifier(string propertyName) {
        if (string.IsNullOrEmpty(propertyName)) return false;

        if (!IsIdentifierStart(propertyName[0])) return false;

        for (int i = 1; i < propertyName.Length; i++) {
            if (!IsIdentifierPart(propertyName[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool IsIdentifierStart(char ch) {
        return ch == '_' || char.IsLetter(ch);
    }

    private static bool IsIdentifierPart(char ch) {
        return ch == '_' || char.IsLetterOrDigit(ch);
    }

    private static string EscapeJsonPathString(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var sb = new StringBuilder(value.Length + 4);
        foreach (var ch in value) {
            switch (ch) {
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '"':
                    sb.Append("\\\"");
                    break;
                case '\b':
                    sb.Append("\\b");
                    break;
                case '\f':
                    sb.Append("\\f");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (char.IsControl(ch)) {
                        sb.Append("\\u");
                        sb.Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        sb.Append(ch);
                    }

                    break;
            }
        }

        return sb.ToString();
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

        return new SourceMetadata {
            Path = path,
            SourceId = sourceId,
            SourceHash = computeHash ? TryComputeFileSha256(path) : null,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static SourceMetadata BuildSourceMetadataFromLogicalStream(Stream stream, string sourcePath, ReaderOptions options, CancellationToken cancellationToken) {
        return new SourceMetadata {
            Path = sourcePath,
            SourceId = BuildSourceId(sourcePath)
        };
    }

    private static void UpdateSourceMetadataFromSeekableStream(SourceMetadata source, Stream stream, bool computeHash) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        try {
            if (stream.CanSeek) {
                source.LengthBytes = stream.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        if (computeHash) {
            source.SourceHash ??= TryComputeStreamSha256(stream);
        }
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

    private sealed class StructuredRow {
        public StructuredRow(string path, string type, string value) {
            Path = path;
            Type = type;
            Value = value;
        }

        public string Path { get; }
        public string Type { get; }
        public string Value { get; }
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
