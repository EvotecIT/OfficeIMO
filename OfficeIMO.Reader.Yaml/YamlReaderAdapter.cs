namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// YAML ingestion helpers for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static class YamlReaderAdapter {
    private static readonly string[] Columns = { "Path", "Type", "Value" };

    /// <summary>
    /// Reads YAML content from a path into path/type/value chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? readerOptions = null, YamlReadOptions? yamlOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in Read(fs, source, effectiveReaderOptions, yamlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads YAML content from a stream into path/type/value chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, YamlReadOptions? yamlOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.yaml");
        return Read(stream, BuildSourceMetadataFromLogicalStream(sourcePath), effectiveReaderOptions, yamlOptions, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> Read(Stream stream, SourceMetadata source, ReaderOptions effectiveReaderOptions, YamlReadOptions? yamlOptions, CancellationToken cancellationToken) {
        int chunkRows = Math.Max(1, yamlOptions?.ChunkRows ?? new YamlReadOptions().ChunkRows);
        bool includeMarkdown = yamlOptions?.IncludeMarkdown ?? true;
        var sourcePath = source.Path;
        bool computeHashes = effectiveReaderOptions.ComputeHashes;

        YamlDocument? document = null;
        string? parseError = null;
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, computeHashes);
            cancellationToken.ThrowIfCancellationRequested();
            document = YamlDocument.Load(parseStream, BuildLoadOptions(yamlOptions?.LoadOptions, effectiveReaderOptions.MaxInputBytes));
        } catch (YamlException ex) {
            parseError = "YAML parse error (" + ex.Code + "): " + ex.Message;
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }

        if (parseError != null) {
            yield return EnrichChunk(BuildWarningChunk(sourcePath, "yaml-warning-0000", parseError), source, computeHashes);
            yield break;
        }

        if (document!.Documents.Count == 0) {
            yield return EnrichChunk(BuildWarningChunk(sourcePath, "yaml-warning-0001", "YAML stream does not contain any documents."), source, computeHashes);
            yield break;
        }

        string[]? loadWarnings = null;
        if (document.Diagnostics.Count > 0) {
            loadWarnings = new string[document.Diagnostics.Count];
            for (int i = 0; i < loadWarnings.Length; i++) {
                var diagnostic = document.Diagnostics[i];
                loadWarnings[i] = "YAML warning (" + diagnostic.Code + "): " + diagnostic.Message;
            }
        }

        var rows = new List<StructuredRow>(Math.Min(chunkRows, 1024));
        var stack = new Stack<PendingNode>();
        bool useDocumentPrefix = document.Documents.Count > 1;
        int chunkIndex = 0;
        int rowIndex = 0;
        for (int documentIndex = 0; documentIndex < document.Documents.Count; documentIndex++) {
            var rootPath = useDocumentPrefix
                ? "$[" + documentIndex.ToString(CultureInfo.InvariantCulture) + "]"
                : "$";
            var root = document.Documents[documentIndex].Root;
            stack.Push(new PendingNode(root, rootPath, root));

            while (stack.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                var pending = stack.Pop();
                AddRows(pending, rows, stack, document.Diagnostics.Count > 0);

                if (rows.Count == chunkRows) {
                    yield return BuildChunk(source, rows, chunkIndex, rowIndex, includeMarkdown, computeHashes, chunkIndex == 0 ? loadWarnings : null);
                    rowIndex += rows.Count;
                    rows.Clear();
                    chunkIndex++;
                }
            }

            // Each document of a multi-document stream starts a new chunk.
            if (rows.Count > 0) {
                yield return BuildChunk(source, rows, chunkIndex, rowIndex, includeMarkdown, computeHashes, chunkIndex == 0 ? loadWarnings : null);
                rowIndex += rows.Count;
                rows.Clear();
                chunkIndex++;
            }
        }
    }

    /// <summary>
    /// Emits the row for a scalar or empty collection, or queues the entries of a collection in document order.
    /// Aliases expand to their anchored node; their rows keep the alias's source lines.
    /// </summary>
    private static void AddRows(PendingNode pending, List<StructuredRow> rows, Stack<PendingNode> stack, bool hasDuplicateKeys) {
        var node = pending.Node is YamlAlias alias ? alias.Target : pending.Node;
        // Inside an alias expansion every row reports the alias's lines rather than the anchored node's.
        bool expanded = pending.Node is YamlAlias || !ReferenceEquals(pending.Location, pending.Node);
        switch (node) {
            case YamlScalar scalar:
                rows.Add(new StructuredRow(pending.Path, Classify(scalar, out var value), value, pending.Location));
                break;
            case YamlSequence sequence:
                if (sequence.Count == 0) {
                    rows.Add(new StructuredRow(pending.Path, "sequence", "[]", pending.Location));
                    break;
                }

                for (int i = sequence.Count - 1; i >= 0; i--) {
                    var item = sequence[i];
                    stack.Push(new PendingNode(
                        item,
                        pending.Path + "[" + i.ToString(CultureInfo.InvariantCulture) + "]",
                        expanded ? pending.Location : item));
                }

                break;
            case YamlMapping mapping:
                if (mapping.Count == 0) {
                    rows.Add(new StructuredRow(pending.Path, "mapping", "{}", pending.Location));
                    break;
                }

                for (int i = mapping.Count - 1; i >= 0; i--) {
                    var entry = mapping[i];
                    var keyText = GetKeyText(entry.Key, out bool scalarKey);
                    // Under a lenient duplicate-key policy only the entry a lookup returns is reported, so every path selects its node.
                    if (hasDuplicateKeys && scalarKey && mapping.TryGetValue(keyText, out var winner) && !ReferenceEquals(winner, entry.Value)) {
                        continue;
                    }

                    stack.Push(new PendingNode(
                        entry.Value,
                        AppendKeyPath(pending.Path, keyText),
                        expanded ? pending.Location : entry.Value));
                }

                break;
        }
    }

    /// <summary>
    /// Returns the exact text a key is addressed by: the decoded value of a scalar key (through an alias), or the
    /// source text of a collection key.
    /// </summary>
    private static string GetKeyText(YamlNode key, out bool scalarKey) {
        var resolved = key is YamlAlias alias ? alias.Target : key;
        if (resolved is YamlScalar scalar) {
            scalarKey = true;
            return scalar.AsString();
        }

        scalarKey = false;
        return resolved.ToString();
    }

    private static string AppendKeyPath(string path, string keyText) {
        if (IsSimplePathIdentifier(keyText)) {
            return path + "." + keyText;
        }

        return path + "[\"" + EscapePathString(keyText) + "\"]";
    }

    /// <summary>
    /// Classifies a scalar through its document's schema (YAML 1.2 core, or YAML 1.1 for <c>%YAML 1.1</c> documents) and
    /// explicit tags, and returns its decoded value.
    /// </summary>
    private static string Classify(YamlScalar scalar, out string value) {
        YamlScalarKind kind = scalar.Kind;
        if (kind == YamlScalarKind.Null) {
            value = "null";
            return "null";
        }

        value = scalar.AsString();
        return kind switch {
            YamlScalarKind.Boolean => "boolean",
            YamlScalarKind.Integer or YamlScalarKind.Float => "number",
            _ => "string"
        };
    }

    private static YamlLoadOptions BuildLoadOptions(YamlLoadOptions? configured, long? maxInputBytes) {
        var options = CopyLoadOptions(configured) ?? new YamlLoadOptions();
        if (maxInputBytes is > 0 && (options.MaxInputBytes is null || options.MaxInputBytes > maxInputBytes)) {
            options.MaxInputBytes = maxInputBytes;
        }

        return options;
    }

    internal static YamlLoadOptions? CopyLoadOptions(YamlLoadOptions? options) => options?.Clone();

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
            startLine = Math.Min(startLine, row.Location.Start.Line);
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
            Id = "yaml-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
            Kind = ReaderInputKind.Yaml,
            Location = new ReaderLocation {
                Path = source.Path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = rowIndex,
                StartLine = startLine,
                EndLine = endLine
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
            Kind = ReaderInputKind.Yaml,
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
        public PendingNode(YamlNode node, string path, YamlNode location) {
            Node = node;
            Path = path;
            Location = location;
        }

        public YamlNode Node { get; }

        public string Path { get; }

        public YamlNode Location { get; }
    }

    private sealed class StructuredRow {
        public StructuredRow(string path, string type, string value, YamlNode location) {
            Path = path;
            Type = type;
            Value = value;
            Location = location;
        }

        public string Path { get; }

        public string Type { get; }

        public string Value { get; }

        public YamlNode Location { get; }
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;

        public string SourceId { get; set; } = string.Empty;

        public string? SourceHash { get; set; }

        public DateTime? LastWriteUtc { get; set; }

        public long? LengthBytes { get; set; }
    }
}
