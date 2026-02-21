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

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadJson(fs, path, effectiveReaderOptions, jsonOptions, cancellationToken)) {
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

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);

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
                yield return BuildWarningChunk(sourcePath, "json-warning-0000", parseError);
                yield break;
            }

            using (doc) {
                var rows = new List<StructuredRow>(capacity: 1024);
                TraverseJson(doc!.RootElement, "$", depth: 0, options.MaxDepth, rows, cancellationToken);

                foreach (var chunk in BuildStructuredChunks(sourcePath, rows, options, cancellationToken)) {
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
        string path,
        IReadOnlyList<StructuredRow> rows,
        JsonReadOptions options,
        CancellationToken cancellationToken) {
        if (rows.Count == 0) {
            yield break;
        }

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

            yield return new ReaderChunk {
                Id = "json-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Text,
                Location = new ReaderLocation {
                    Path = path,
                    BlockIndex = chunkIndex,
                    SourceBlockIndex = index
                },
                Text = BuildPlain(slice),
                Markdown = options.IncludeMarkdown ? BuildMarkdown(slice) : null,
                Tables = new[] { table }
            };

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
                    var childPath = path == "$"
                        ? "$." + property.Name
                        : path + "." + property.Name;
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
            Kind = ReaderInputKind.Unknown,
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
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName!;
        }

        return defaultName;
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
}
