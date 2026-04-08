namespace OfficeIMO.Reader.Xml;

/// <summary>
/// XML ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderXmlExtensions {
    /// <summary>
    /// Reads XML content from a path with tree-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadXml(string path, ReaderOptions? readerOptions = null, XmlReadOptions? xmlOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(path, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadXml(fs, source, effectiveReaderOptions, xmlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads XML content from a stream with tree-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadXml(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, XmlReadOptions? xmlOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var options = Normalize(xmlOptions);
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.xml");
        return ReadXml(stream, new SourceMetadata { Path = sourcePath, SourceId = BuildSourceId(sourcePath) }, effectiveReaderOptions, options, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadXml(Stream stream, SourceMetadata source, ReaderOptions effectiveReaderOptions, XmlReadOptions? xmlOptions, CancellationToken cancellationToken) {
        var options = Normalize(xmlOptions);
        var sourcePath = source.Path;
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);
        UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);

        try {
            XDocument? doc = null;
            string? parseError = null;
            try {
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Ignore,
                    XmlResolver = null
                };

                using var reader = XmlReader.Create(parseStream, settings);
                doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                parseError = "XML parse error: " + ex.GetType().Name + ".";
            }

            if (parseError != null) {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "xml-warning-0000", parseError), source, effectiveReaderOptions.ComputeHashes);
                yield break;
            }

            var root = doc!.Root;
            if (root == null) {
                yield return EnrichChunk(BuildWarningChunk(sourcePath, "xml-warning-0001", "XML document does not contain a root element."), source, effectiveReaderOptions.ComputeHashes);
                yield break;
            }

            var rows = new List<StructuredRow>(capacity: 1024);
            TraverseXml(root, parentPath: string.Empty, rows, cancellationToken);

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
        XmlReadOptions options,
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
                Id = "xml-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Xml,
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

    private static void TraverseXml(XElement element, string parentPath, List<StructuredRow> rows, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        var currentPath = BuildXmlPath(element, parentPath);
        rows.Add(new StructuredRow(currentPath, "element", NormalizeText(GetDirectText(element))));

        foreach (var attribute in element.Attributes()) {
            rows.Add(new StructuredRow(
                currentPath + "/@" + GetQualifiedName(element, attribute.Name),
                "attribute",
                NormalizeText(attribute.Value)));
        }

        foreach (var child in element.Elements()) {
            TraverseXml(child, currentPath, rows, cancellationToken);
        }
    }

    private static string BuildXmlPath(XElement element, string parentPath) {
        var siblingIndex = 1 + element.ElementsBeforeSelf(element.Name).Count();
        var segment = GetQualifiedName(element, element.Name) + "[" + siblingIndex.ToString(CultureInfo.InvariantCulture) + "]";
        return parentPath.Length == 0 ? segment : parentPath + "/" + segment;
    }

    private static string GetQualifiedName(XElement context, XName name) {
        if (name.Namespace == XNamespace.None) {
            return name.LocalName;
        }

        var prefix = context.GetPrefixOfNamespace(name.Namespace);
        if (!string.IsNullOrWhiteSpace(prefix)) {
            return prefix + ":" + name.LocalName;
        }

        return "{" + name.NamespaceName + "}" + name.LocalName;
    }

    private static string GetDirectText(XElement element) {
        var sb = new StringBuilder();
        foreach (var node in element.Nodes()) {
            switch (node) {
                case XCData cdata:
                    sb.Append(cdata.Value);
                    sb.Append(' ');
                    break;
                case XText text:
                    sb.Append(text.Value);
                    sb.Append(' ');
                    break;
            }
        }

        return sb.ToString();
    }

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Xml,
            Location = new ReaderLocation { Path = path },
            Text = warning,
            Warnings = new[] { warning }
        };
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

    private static XmlReadOptions Normalize(XmlReadOptions? options) {
        var source = options ?? new XmlReadOptions();

        var normalized = new XmlReadOptions {
            ChunkRows = source.ChunkRows,
            IncludeMarkdown = source.IncludeMarkdown
        };

        if (normalized.ChunkRows < 1) normalized.ChunkRows = 1;

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

        return new SourceMetadata {
            Path = path,
            SourceId = sourceId,
            SourceHash = computeHash ? TryComputeFileSha256(path) : null,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
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
