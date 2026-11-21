namespace OfficeIMO.Reader.Toml;

/// <summary>
/// TOML ingestion helpers for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static class TomlReaderAdapter {
    private static readonly string[] Columns = { "Path", "Type", "Value" };

    /// <summary>
    /// Reads TOML content from a path into path/type/value chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? readerOptions = null, TomlReadOptions? tomlOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in Read(fs, source, effectiveReaderOptions, tomlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads TOML content from a stream into path/type/value chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, TomlReadOptions? tomlOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.toml");
        return Read(stream, BuildSourceMetadataFromLogicalStream(sourcePath), effectiveReaderOptions, tomlOptions, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> Read(Stream stream, SourceMetadata source, ReaderOptions effectiveReaderOptions, TomlReadOptions? tomlOptions, CancellationToken cancellationToken) {
        int chunkRows = Math.Max(1, tomlOptions?.ChunkRows ?? new TomlReadOptions().ChunkRows);
        bool includeMarkdown = tomlOptions?.IncludeMarkdown ?? true;
        bool computeHashes = effectiveReaderOptions.ComputeHashes;
        TomlLoadOptions loadOptions = BuildLoadOptions(tomlOptions, effectiveReaderOptions.MaxInputBytes);

        TomlDocument? document = null;
        string? parseError = null;
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, computeHashes);
            cancellationToken.ThrowIfCancellationRequested();
            document = TomlDocument.Load(parseStream, loadOptions);
        } catch (TomlLimitExceededException ex) {
            throw new ReaderResourceLimitException(ReaderLimitName(ex.LimitName), ex.Limit, ex);
        } catch (TomlException ex) {
            parseError = "TOML parse error (" + ex.Code + "): " + ex.Message;
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }

        if (parseError != null) {
            yield return EnrichChunk(BuildWarningChunk(source.Path, "toml-warning-0000", parseError), source, computeHashes);
            yield break;
        }

        var rows = new List<StructuredRow>(Math.Min(chunkRows, 1024));
        var stack = new Stack<PendingNode>();
        stack.Push(new PendingNode(document!.Root, "$"));
        int chunkIndex = 0;
        int rowIndex = 0;
        while (stack.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            AddRows(stack.Pop(), rows, stack);
            if (rows.Count == chunkRows) {
                yield return BuildChunk(source, rows, chunkIndex, rowIndex, includeMarkdown, computeHashes, null);
                rowIndex += rows.Count;
                rows.Clear();
                chunkIndex++;
            }
        }

        if (rows.Count > 0 || chunkIndex == 0) {
            if (rows.Count == 0) rows.Add(new StructuredRow("$", "table", "{}", document.Root));
            yield return BuildChunk(source, rows, chunkIndex, rowIndex, includeMarkdown, computeHashes, null);
        }
    }

    /// <summary>
    /// Emits the row for a value or empty container, or queues the entries of a container in document order.
    /// </summary>
    private static void AddRows(PendingNode pending, List<StructuredRow> rows, Stack<PendingNode> stack) {
        switch (pending.Node) {
            case TomlValue value:
                rows.Add(new StructuredRow(pending.Path, TypeName(value), value.AsString(), value));
                break;
            case TomlArray array:
                if (array.Count == 0) {
                    rows.Add(new StructuredRow(pending.Path, "array", "[]", array));
                    break;
                }
                for (int i = array.Count - 1; i >= 0; i--) {
                    stack.Push(new PendingNode(array[i], pending.Path + "[" + i.ToString(CultureInfo.InvariantCulture) + "]"));
                }
                break;
            case TomlTableArray tables:
                for (int i = tables.Count - 1; i >= 0; i--) {
                    stack.Push(new PendingNode(tables[i], pending.Path + "[" + i.ToString(CultureInfo.InvariantCulture) + "]"));
                }
                break;
            case TomlTable table:
                if (table.Count == 0) {
                    if (pending.Path != "$") rows.Add(new StructuredRow(pending.Path, "table", "{}", table));
                    break;
                }
                var entries = table.ToList();
                for (int i = entries.Count - 1; i >= 0; i--) {
                    stack.Push(new PendingNode(entries[i].Value, AppendKeyPath(pending.Path, entries[i].Key)));
                }
                break;
        }
    }

    private static string TypeName(TomlValue value) {
        switch (value.Kind) {
            case TomlNodeKind.String: return "string";
            case TomlNodeKind.Integer: return "integer";
            case TomlNodeKind.Float: return "float";
            case TomlNodeKind.Boolean: return "boolean";
        }
        return ((TomlDateTime)value.Value).Kind switch {
            TomlDateTimeKind.OffsetDateTime => "datetime",
            TomlDateTimeKind.LocalDateTime => "datetime-local",
            TomlDateTimeKind.LocalDate => "date-local",
            _ => "time-local"
        };
    }

    private static string AppendKeyPath(string path, string keyText) {
        if (IsSimplePathIdentifier(keyText)) {
            return path + "." + keyText;
        }

        return path + "[\"" + EscapePathString(keyText) + "\"]";
    }

    private static TomlLoadOptions BuildLoadOptions(TomlReadOptions? readOptions, long? maxInputBytes) {
        var options = CopyLoadOptions(readOptions?.LoadOptions) ?? TomlLoadOptions.UntrustedDefaults;
        if (readOptions != null) {
            if (readOptions.MaxDepth < 1 || readOptions.MaxNodes < 1 || readOptions.MaxScalarLength < 1) {
                throw new ArgumentOutOfRangeException(nameof(readOptions), "TOML parser limits must be positive.");
            }

            if (readOptions.MaxDepth is int maxDepth) options.MaxDepth = maxDepth;
            if (readOptions.MaxNodes is int maxNodes) options.MaxNodes = maxNodes;
            if (readOptions.MaxScalarLength is int maxScalarLength) options.MaxStringLength = maxScalarLength;
        }

        if (maxInputBytes is > 0 && (options.MaxInputBytes is null || options.MaxInputBytes > maxInputBytes)) {
            options.MaxInputBytes = maxInputBytes;
        }

        return options;
    }

    // The engine names its string limit MaxStringLength; Reader callers configure it as MaxScalarLength.
    private static string ReaderLimitName(string engineLimitName) =>
        engineLimitName == nameof(TomlLoadOptions.MaxStringLength) ? nameof(TomlReadOptions.MaxScalarLength) : engineLimitName;

    internal static TomlLoadOptions? CopyLoadOptions(TomlLoadOptions? options) {
        if (options == null) return null;

        return new TomlLoadOptions {
            MaxDepth = options.MaxDepth,
            MaxNodes = options.MaxNodes,
            MaxStringLength = options.MaxStringLength,
            MaxInputBytes = options.MaxInputBytes
        };
    }

    private static ReaderChunk BuildChunk(
        SourceMetadata source,
        List<StructuredRow> rows,
        int chunkIndex,
        int rowIndex,
        bool includeMarkdown,
        bool computeHashes,
        string[]? warnings) {
        var tableRows = new IReadOnlyList<string>[rows.Count];
        int startLine = int.MaxValue;
        int endLine = 0;
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            tableRows[i] = new[] { row.Path, row.Type, row.Value };
            // Values added by editing have no line; loaded rows always do.
            if (row.Location.Start.Line > 0) startLine = Math.Min(startLine, row.Location.Start.Line);
            endLine = Math.Max(endLine, row.Location.End.Line);
        }

        var table = new ReaderTable {
            Title = Path.GetFileName(source.Path),
            Columns = Columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(Columns, tableRows),
            Rows = tableRows,
            TotalRowCount = rows.Count,
            Truncated = false
        };

        return EnrichChunk(new ReaderChunk {
            Id = "toml-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
            Kind = ReaderInputKind.Toml,
            Location = new ReaderLocation {
                Path = source.Path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = rowIndex,
                StartLine = startLine == int.MaxValue ? null : startLine,
                EndLine = endLine == 0 ? null : endLine
            },
            Text = BuildPlain(rows),
            Markdown = includeMarkdown ? BuildMarkdown(rows) : null,
            Tables = new[] { table },
            Warnings = warnings
        }, source, computeHashes);
    }

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Toml,
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
            sb.AppendLine(EscapePlainValue(row.Value));
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
        return EscapePlainValue(value)
            .Replace("\\", "\\\\")
            .Replace("|", "\\|");
    }

    private static string EscapePlainValue(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var sb = new StringBuilder(value.Length + 8);
        foreach (var ch in value) {
            switch (ch) {
                case '\r':
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }

        return sb.ToString();
    }

    private static bool IsSimplePathIdentifier(string value) {
        if (string.IsNullOrEmpty(value)) return false;

        if (!IsIdentifierStart(value[0])) return false;

        for (int i = 1; i < value.Length; i++) {
            if (!IsIdentifierPart(value[i])) {
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

    private static string EscapePathString(string value) {
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

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                return trimmedSourceName;
            }
        }

        return defaultName;
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

    private static SourceMetadata BuildSourceMetadataFromLogicalStream(string sourcePath) {
        return new SourceMetadata {
            Path = sourcePath,
            SourceId = BuildSourceId(sourcePath)
        };
    }

    private static void UpdateSourceMetadataFromSeekableStream(SourceMetadata source, Stream stream, bool computeHash) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(stream);

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
        ArgumentNullException.ThrowIfNull(chunk);
        ArgumentNullException.ThrowIfNull(source);

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

    private static string ComputeChunkHash(ReaderChunk chunk) => DocumentReaderEngine.ComputeChunkHash(chunk);

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

    private readonly struct PendingNode {
        public PendingNode(TomlNode node, string path) {
            Node = node;
            Path = path;
        }

        public TomlNode Node { get; }

        public string Path { get; }
    }

    private sealed class StructuredRow {
        public StructuredRow(string path, string type, string value, TomlNode location) {
            Path = path;
            Type = type;
            Value = value;
            Location = location;
        }

        public string Path { get; }

        public string Type { get; }

        public string Value { get; }

        public TomlNode Location { get; }
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;

        public string SourceId { get; set; } = string.Empty;

        public string? SourceHash { get; set; }

        public DateTime? LastWriteUtc { get; set; }

        public long? LengthBytes { get; set; }
    }
}
