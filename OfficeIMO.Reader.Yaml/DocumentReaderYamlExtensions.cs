namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// YAML ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderYamlExtensions {
    /// <summary>
    /// Reads YAML content from a path with representation-model-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadYaml(string path, ReaderOptions? readerOptions = null, YamlReadOptions? yamlOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadYaml(fs, source, effectiveReaderOptions, yamlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads YAML content from a stream with representation-model-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadYaml(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, YamlReadOptions? yamlOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var options = Normalize(yamlOptions);
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.yaml");
        return ReadYaml(stream, BuildSourceMetadataFromLogicalStream(sourcePath), effectiveReaderOptions, options, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadYaml(Stream stream, SourceMetadata source, ReaderOptions effectiveReaderOptions, YamlReadOptions? yamlOptions, CancellationToken cancellationToken) {
        var options = Normalize(yamlOptions);
        var sourcePath = source.Path;

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);
        UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);

        YamlStream? yaml = null;
        string? parseError = null;
        try {
            try {
                using var reader = new StreamReader(parseStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
                yaml = new YamlStream();
                yaml.Load(reader);
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                parseError = "YAML parse error: " + ex.GetType().Name + ".";
            }

            if (parseError != null) {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "yaml-warning-0000", parseError), source, effectiveReaderOptions.ComputeHashes);
                yield break;
            }

            if (yaml == null || yaml.Documents.Count == 0) {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "yaml-warning-0001", "YAML stream does not contain any documents."), source, effectiveReaderOptions.ComputeHashes);
                yield break;
            }

            var rows = new List<StructuredRow>(capacity: 1024);
            var state = new TraversalState();
            var useDocumentPrefix = yaml.Documents.Count > 1;
            for (int documentIndex = 0; documentIndex < yaml.Documents.Count; documentIndex++) {
                cancellationToken.ThrowIfCancellationRequested();

                var document = yaml.Documents[documentIndex];
                var rootPath = useDocumentPrefix
                    ? "$[" + documentIndex.ToString(CultureInfo.InvariantCulture) + "]"
                    : "$";

                if (document.RootNode == null) {
                    rows.Add(new StructuredRow(rootPath, "null", "null"));
                    continue;
                }

                TraverseYaml(document.RootNode, rootPath, depth: 0, options, rows, state, cancellationToken);
                if (state.NodeLimitEmitted) {
                    break;
                }
            }

            foreach (var chunk in BuildStructuredChunks(source, rows, options, effectiveReaderOptions.ComputeHashes, cancellationToken)) {
                yield return chunk;
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
        YamlReadOptions options,
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
            var columns = new[] { "Path", "Type", "Value" };

            var table = new ReaderTable {
                Title = Path.GetFileName(path),
                Columns = columns,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, tableRows),
                Rows = tableRows,
                TotalRowCount = slice.Count,
                Truncated = false
            };

            yield return EnrichChunk(new ReaderChunk {
                Id = "yaml-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Yaml,
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

    private static void TraverseYaml(
        YamlNode node,
        string path,
        int depth,
        YamlReadOptions options,
        List<StructuredRow> rows,
        TraversalState state,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        if (state.NodeLimitEmitted) {
            return;
        }

        if (depth > options.MaxDepth) {
            rows.Add(new StructuredRow(path, "depth-limit", "(max depth reached)"));
            return;
        }

        if (!TryVisitNode(path, options, rows, state)) {
            return;
        }

        switch (node) {
            case YamlScalarNode scalar:
                rows.Add(new StructuredRow(path, ClassifyScalar(scalar), NormalizeScalarValue(scalar)));
                break;
            case YamlSequenceNode sequence:
                if (sequence.Children.Count == 0) {
                    rows.Add(new StructuredRow(path, "sequence", "[]"));
                    break;
                }

                for (int i = 0; i < sequence.Children.Count; i++) {
                    TraverseYaml(
                        sequence.Children[i],
                        path + "[" + i.ToString(CultureInfo.InvariantCulture) + "]",
                        depth + 1,
                        options,
                        rows,
                        state,
                        cancellationToken);

                    if (state.NodeLimitEmitted) {
                        break;
                    }
                }

                break;
            case YamlMappingNode mapping:
                if (mapping.Children.Count == 0) {
                    rows.Add(new StructuredRow(path, "mapping", "{}"));
                    break;
                }

                foreach (var child in mapping.Children) {
                    TraverseYaml(
                        child.Value,
                        AppendKeyPath(path, child.Key),
                        depth + 1,
                        options,
                        rows,
                        state,
                        cancellationToken);

                    if (state.NodeLimitEmitted) {
                        break;
                    }
                }

                break;
            default:
                rows.Add(new StructuredRow(path, node.NodeType.ToString().ToLowerInvariant(), NormalizeText(node.ToString())));
                break;
        }
    }

    private static bool TryVisitNode(string path, YamlReadOptions options, List<StructuredRow> rows, TraversalState state) {
        if (state.NodesVisited >= options.MaxNodes) {
            if (!state.NodeLimitEmitted) {
                rows.Add(new StructuredRow(path, "node-limit", "(max nodes reached)"));
                state.NodeLimitEmitted = true;
            }

            return false;
        }

        state.NodesVisited++;
        return true;
    }

    private static string AppendKeyPath(string path, YamlNode key) {
        var keyText = GetKeyText(key);
        if (IsSimplePathIdentifier(keyText)) {
            return path + "." + keyText;
        }

        return path + "[\"" + EscapePathString(keyText) + "\"]";
    }

    private static string GetKeyText(YamlNode key) {
        if (key is YamlScalarNode scalar) {
            return scalar.Value ?? string.Empty;
        }

        return NormalizeText(key.ToString());
    }

    private static string ClassifyScalar(YamlScalarNode scalar) {
        var value = scalar.Value;
        if (scalar.Style == ScalarStyle.Plain && !HasExplicitStringTag(scalar)) {
            if (IsYamlNull(value)) {
                return "null";
            }

            if (IsYamlBoolean(value)) {
                return "boolean";
            }

            if (IsYamlNumber(value)) {
                return "number";
            }
        }

        return "string";
    }

    private static string NormalizeScalarValue(YamlScalarNode scalar) {
        var value = scalar.Value;
        if (scalar.Style == ScalarStyle.Plain && !HasExplicitStringTag(scalar) && IsYamlNull(value)) {
            return "null";
        }

        if (scalar.Style == ScalarStyle.Literal || scalar.Style == ScalarStyle.Folded) {
            return NormalizeBlockScalarValue(value ?? string.Empty);
        }

        return NormalizeText(value ?? string.Empty);
    }

    private static bool HasExplicitStringTag(YamlScalarNode scalar) {
        var tag = scalar.Tag.ToString();
        return string.Equals(tag, "tag:yaml.org,2002:str", StringComparison.Ordinal) ||
               string.Equals(tag, "!!str", StringComparison.Ordinal);
    }

    private static string NormalizeBlockScalarValue(string value) {
        if (value.Length == 0) return string.Empty;

        var normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Length > 2048) {
            normalized = normalized.Substring(0, 2048);
        }

        return normalized;
    }

    private static bool IsYamlNull(string? value) {
        if (value == null) return true;

        var trimmed = value!.Trim();
        return trimmed.Length == 0 ||
               string.Equals(trimmed, "~", StringComparison.Ordinal) ||
               string.Equals(trimmed, "null", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsYamlBoolean(string? value) {
        if (value == null) return false;

        var trimmed = value!.Trim();
        return string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "false", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsYamlNumber(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;

        var trimmed = value!.Trim();
        if (trimmed.Equals(".", StringComparison.Ordinal) ||
            trimmed.Equals("+", StringComparison.Ordinal) ||
            trimmed.Equals("-", StringComparison.Ordinal)) {
            return false;
        }

        return double.TryParse(
            trimmed,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out var number) &&
            !double.IsNaN(number) &&
            !double.IsInfinity(number);
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

    private static string NormalizeText(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var sb = new StringBuilder(value.Length);
        bool previousWhitespace = false;
        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWhitespace) {
                    sb.Append(' ');
                    previousWhitespace = true;
                }
            } else {
                sb.Append(ch);
                previousWhitespace = false;
            }
        }

        var normalized = sb.ToString().Trim();
        if (normalized.Length > 2048) {
            normalized = normalized.Substring(0, 2048);
        }

        return normalized;
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

    private static YamlReadOptions Normalize(YamlReadOptions? options) {
        var source = options ?? new YamlReadOptions();

        var normalized = new YamlReadOptions {
            ChunkRows = source.ChunkRows,
            MaxDepth = source.MaxDepth,
            MaxNodes = source.MaxNodes,
            IncludeMarkdown = source.IncludeMarkdown
        };

        if (normalized.ChunkRows < 1) normalized.ChunkRows = 1;
        if (normalized.MaxDepth < 1) normalized.MaxDepth = 1;
        if (normalized.MaxNodes < 1) normalized.MaxNodes = 1;

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

    private sealed class TraversalState {
        public int NodesVisited { get; set; }

        public bool NodeLimitEmitted { get; set; }
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
