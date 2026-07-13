using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

/// <summary>
/// Visio ingestion adapter for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static partial class VisioReaderAdapter {
    /// <summary>
    /// Reads a Visio file and emits normalized page chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioPath == null) throw new ArgumentNullException(nameof(visioPath));
        if (visioPath.Length == 0) throw new ArgumentException("Visio path cannot be empty.", nameof(visioPath));
        if (!File.Exists(visioPath)) throw new FileNotFoundException($"Visio file '{visioPath}' doesn't exist.", visioPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(visioPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(visioPath, effectiveReaderOptions.ComputeHashes);
        VisioDocument document = VisioDocument.Load(visioPath);
        foreach (var chunk in Read(document, source, effectiveReaderOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads a Visio stream and emits normalized page chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioStream == null) throw new ArgumentNullException(nameof(visioStream));
        if (!visioStream.CanRead) throw new ArgumentException("Visio stream must be readable.", nameof(visioStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.vsdx");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        Stream parseStream = EnsureParseStream(visioStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);
            if (parseStream.CanSeek) {
                parseStream.Position = 0;
            }

            VisioDocument document = VisioDocument.Load(parseStream);
            foreach (var chunk in Read(document, source, effectiveReaderOptions, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Reads an already loaded Visio document and emits normalized page chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(VisioDocument document, string sourceName = "document.vsdx", ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.vsdx");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        foreach (var chunk in Read(document, source, effectiveReaderOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> Read(VisioDocument document, SourceMetadata source, ReaderOptions readerOptions, CancellationToken cancellationToken) {
        VisioInspectionSnapshot snapshot = document.CreateInspectionSnapshot();
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            string markdown = BuildPageMarkdown(snapshot, page);
            string text = BuildPageText(page);
            ReaderTable? shapeDataTable = BuildShapeDataTable(page, source, pageIndex, readerOptions.MaxTableRows);

            foreach (ReaderChunk chunk in BuildPageChunks(source, pageIndex, text, markdown, shapeDataTable, readerOptions.MaxChars)) {
                yield return EnrichChunk(chunk, source, readerOptions.ComputeHashes);
            }
        }
    }

    private static IEnumerable<ReaderChunk> BuildPageChunks(SourceMetadata source, int pageIndex, string text, string markdown, ReaderTable? shapeDataTable, int maxChars) {
        maxChars = Math.Max(1, maxChars);
        if (text.Length <= maxChars && markdown.Length <= maxChars) {
            yield return new ReaderChunk {
                Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Visio,
                Location = BuildLocation(source, pageIndex, "page", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture)),
                Text = text,
                Markdown = markdown,
                Tables = shapeDataTable == null ? null : new[] { shapeDataTable }
            };
            yield break;
        }

        IReadOnlyList<string> textParts = SplitByMaxChars(text, maxChars);
        IReadOnlyList<string> markdownParts = SplitByMaxChars(markdown, maxChars);
        int partCount = Math.Max(textParts.Count, markdownParts.Count);
        for (int partIndex = 0; partIndex < partCount; partIndex++) {
            string blockAnchor = "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-part-" + (partIndex + 1).ToString(CultureInfo.InvariantCulture);
            yield return new ReaderChunk {
                Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-part-" + (partIndex + 1).ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Visio,
                Location = BuildLocation(source, pageIndex, "page", blockAnchor),
                Text = partIndex < textParts.Count ? textParts[partIndex] : string.Empty,
                Markdown = partIndex < markdownParts.Count ? markdownParts[partIndex] : string.Empty,
                Tables = partIndex == 0 && shapeDataTable != null ? new[] { shapeDataTable } : null
            };
        }
    }

    private static IReadOnlyList<string> SplitByMaxChars(string value, int maxChars) {
        if (string.IsNullOrEmpty(value)) {
            return new[] { string.Empty };
        }

        var parts = new List<string>();
        int offset = 0;
        while (offset < value.Length) {
            int length = Math.Min(maxChars, value.Length - offset);
            int cut = length;
            if (offset + length < value.Length) {
                int newlineIndex = value.LastIndexOf('\n', offset + length - 1, length);
                if (newlineIndex > offset) {
                    cut = newlineIndex - offset + 1;
                } else {
                    int spaceIndex = value.LastIndexOf(' ', offset + length - 1, length);
                    if (spaceIndex > offset) {
                        cut = spaceIndex - offset + 1;
                    }
                }
            }

            string part = value.Substring(offset, cut).Trim();
            if (part.Length == 0) {
                part = value.Substring(offset, length);
                cut = length;
            }

            parts.Add(part);
            offset += cut;
            while (offset < value.Length && char.IsWhiteSpace(value[offset])) {
                offset++;
            }
        }

        return parts;
    }

    internal static string BuildPageMarkdown(VisioInspectionSnapshot snapshot, VisioInspectionPageSnapshot page) {
        var builder = new StringBuilder();
        builder.Append("# ");
        builder.AppendLine(string.IsNullOrWhiteSpace(page.Name) ? "Page " + page.Id.ToString(CultureInfo.InvariantCulture) : page.Name);
        builder.AppendLine();
        builder.Append("Document: ");
        builder.AppendLine(string.IsNullOrWhiteSpace(snapshot.Title) ? "Untitled Visio document" : snapshot.Title);
        builder.Append("- Shapes: ");
        builder.AppendLine(page.Shapes.Count.ToString(CultureInfo.InvariantCulture));
        builder.Append("- Connectors: ");
        builder.AppendLine(page.Connectors.Count.ToString(CultureInfo.InvariantCulture));
        if (page.Layers.Count > 0) {
            builder.Append("- Layers: ");
            builder.AppendLine(string.Join(", ", page.Layers));
        }

        AppendShapeMarkdown(builder, page);
        AppendConnectorMarkdown(builder, page);
        return builder.ToString().TrimEnd();
    }

    private static string BuildPageText(VisioInspectionPageSnapshot page) {
        var parts = new List<string> {
            "Visio page " + page.Name + ": " + page.Shapes.Count.ToString(CultureInfo.InvariantCulture) + " shape(s), " + page.Connectors.Count.ToString(CultureInfo.InvariantCulture) + " connector(s)."
        };
        parts.AddRange(page.Shapes.Select(shape => string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text!));
        parts.AddRange(page.Connectors.Select(connector => string.IsNullOrWhiteSpace(connector.Label) ? connector.FromId + " -> " + connector.ToId : connector.Label!));
        return string.Join(Environment.NewLine, parts);
    }

    private static void AppendShapeMarkdown(StringBuilder builder, VisioInspectionPageSnapshot page) {
        if (page.Shapes.Count == 0) return;

        builder.AppendLine();
        builder.AppendLine("## Shapes");
        foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
            builder.Append("- ");
            builder.Append(string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text);
            builder.Append(" (`");
            builder.Append(shape.Id);
            builder.Append('`');
            if (!string.IsNullOrWhiteSpace(shape.MasterNameU)) {
                builder.Append(", master `");
                builder.Append(shape.MasterNameU);
                builder.Append('`');
            }
            builder.Append(')');
            if (shape.ShapeData.Count > 0) {
                builder.Append(": ");
                builder.Append(string.Join("; ", shape.ShapeData.Select(FormatShapeData)));
            }
            builder.AppendLine();
        }
    }

    private static void AppendConnectorMarkdown(StringBuilder builder, VisioInspectionPageSnapshot page) {
        if (page.Connectors.Count == 0) return;

        builder.AppendLine();
        builder.AppendLine("## Connectors");
        foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
            builder.Append("- ");
            builder.Append(connector.FromId);
            builder.Append(" -> ");
            builder.Append(connector.ToId);
            if (!string.IsNullOrWhiteSpace(connector.Label)) {
                builder.Append(": ");
                builder.Append(connector.Label);
            }
            if (connector.ShapeData.Count > 0) {
                builder.Append(" (");
                builder.Append(string.Join("; ", connector.ShapeData.Select(FormatShapeData)));
                builder.Append(')');
            }
            builder.AppendLine();
        }
    }

    internal static ReaderTable? BuildShapeDataTable(VisioInspectionPageSnapshot page, SourceMetadata source, int pageIndex, int maxTableRows) {
        var rows = new List<IReadOnlyList<string>>();
        foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
            AddShapeDataRows(rows, "shape", shape.Id, shape.Text, shape.ShapeData);
        }
        foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
            AddShapeDataRows(rows, "connector", connector.Id, connector.Label, connector.ShapeData);
        }

        if (rows.Count == 0) {
            return null;
        }

        string[] columns = { "OwnerType", "OwnerId", "OwnerText", "Name", "Label", "Value", "Type", "Prompt" };
        int totalRowCount = rows.Count;
        IReadOnlyList<IReadOnlyList<string>> visibleRows = rows.Count > maxTableRows
            ? rows.Take(maxTableRows).ToArray()
            : rows;
        return new ReaderTable {
            Title = page.Name + " Shape Data",
            Kind = "visio-shape-data",
            Location = BuildLocation(source, pageIndex, "shape-data", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-shape-data"),
            Columns = columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, visibleRows),
            Rows = visibleRows,
            TotalRowCount = totalRowCount,
            Truncated = totalRowCount > visibleRows.Count
        };
    }

    private static void AddShapeDataRows(List<IReadOnlyList<string>> rows, string ownerType, string ownerId, string? ownerText, IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeDataRows) {
        foreach (VisioInspectionShapeDataSnapshot row in shapeDataRows) {
            rows.Add(new[] {
                ownerType,
                ownerId,
                ownerText ?? string.Empty,
                row.Name,
                row.Label ?? string.Empty,
                row.Value ?? string.Empty,
                row.Type ?? string.Empty,
                row.Prompt ?? string.Empty
            });
        }
    }

    private static string FormatShapeData(VisioInspectionShapeDataSnapshot row) {
        string name = string.IsNullOrWhiteSpace(row.Label) ? row.Name : row.Label!;
        return name + "=" + (row.Value ?? string.Empty);
    }

    internal static ReaderLocation BuildLocation(SourceMetadata source, int pageIndex, string sourceBlockKind, string blockAnchor) {
        return new ReaderLocation {
            Path = source.Path,
            Page = pageIndex + 1,
            SourceBlockIndex = pageIndex,
            SourceBlockKind = sourceBlockKind,
            BlockAnchor = blockAnchor
        };
    }

    internal static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceMetadata source, bool computeHashes) {
        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeSha256Hex(string.Join("|",
                chunk.Kind.ToString(),
                chunk.SourceId ?? string.Empty,
                chunk.Location.Path ?? string.Empty,
                chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                chunk.Location.BlockAnchor ?? string.Empty,
                chunk.Text ?? string.Empty,
                chunk.Markdown ?? string.Empty));
        }
        return chunk;
    }

    internal static int EstimateTokenCount(string? text) {
        string safeText = text ?? string.Empty;
        return safeText.Length == 0 ? 0 : Math.Max(1, (safeText.Length + 3) / 4);
    }

    internal static SourceMetadata BuildSourceMetadataFromPath(string path, bool computeHash) {
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
            SourceId = BuildSourceId(NormalizePathForId(path)),
            SourceHash = computeHash ? TryComputeFileSha256(path) : null,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    internal static void UpdateSourceMetadataFromSeekableStream(SourceMetadata source, Stream stream, bool computeHash) {
        if (stream.CanSeek) {
            source.LengthBytes = stream.Length;
        }

        if (computeHash) {
            source.SourceHash = TryComputeStreamSha256(stream);
        }
    }

    internal static Stream EnsureParseStream(Stream stream, long? maxInputBytes, CancellationToken cancellationToken, out bool ownsStream) {
        if (stream.CanSeek) {
            ReaderInputLimits.EnforceSeekableStreamSize(stream, maxInputBytes);
        }

        var buffer = new MemoryStream();
        try {
            var chunk = new byte[64 * 1024];
            long totalBytes = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(chunk, 0, chunk.Length);
                if (read <= 0) break;
                buffer.Write(chunk, 0, read);
                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        "Input exceeds MaxInputBytes (" + totalBytes.ToString(CultureInfo.InvariantCulture) + " > " + maxInputBytes.Value.ToString(CultureInfo.InvariantCulture) + ").");
                }
            }
        } catch {
            buffer.Dispose();
            throw;
        }

        buffer.Position = 0;
        ownsStream = true;
        return buffer;
    }

    internal static string NormalizeLogicalSourceName(string? sourceName, string fallback) {
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName!.Trim();
        }

        return fallback;
    }

    internal static string BuildSourceId(string sourceKey) {
        string normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar == '\\') {
            normalized = normalized.ToLowerInvariant();
        }
        return "src:" + ComputeSha256Hex(normalized);
    }

    internal static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        try {
            return Path.GetFullPath(path).Replace('\\', '/');
        } catch {
            return path.Replace('\\', '/');
        }
    }

    internal static string? TryComputeFileSha256(string path) {
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(stream);
        } catch {
            return null;
        }
    }

    internal static string? TryComputeStreamSha256(Stream stream) {
        if (stream == null || !stream.CanSeek) return null;
        long position = stream.Position;
        try {
            stream.Position = 0;
            string hash = ComputeSha256Hex(stream);
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

    internal static string ComputeSha256Hex(string value) {
        return ComputeSha256Hex(Encoding.UTF8.GetBytes(value ?? string.Empty));
    }

    internal static string ComputeSha256Hex(Stream stream) {
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(stream));
    }

    internal static string ComputeSha256Hex(byte[] bytes) {
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(bytes));
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }
        return builder.ToString();
    }

    internal sealed class SourceMetadata {
        public string? Path { get; set; }
        public string? SourceId { get; set; }
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
