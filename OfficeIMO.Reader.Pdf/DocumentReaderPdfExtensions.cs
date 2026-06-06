using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// PDF ingestion adapter for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderPdfExtensions {
    /// <summary>
    /// Reads a PDF file and emits normalized page-aware chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadPdfFile(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));
        if (pdfPath.Length == 0) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));
        if (!File.Exists(pdfPath)) throw new FileNotFoundException($"PDF file '{pdfPath}' doesn't exist.", pdfPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        ReaderInputLimits.EnforceFileSize(pdfPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(pdfPath, effectiveReaderOptions.ComputeHashes);

        PdfLogicalDocument document = LoadDocument(pdfPath, effectivePdfOptions);
        foreach (var chunk in ReadPdf(document, source, effectiveReaderOptions, effectivePdfOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads a PDF stream and emits normalized page-aware chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadPdf(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        if (!pdfStream.CanRead) throw new ArgumentException("PDF stream must be readable.", nameof(pdfStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(pdfStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out var ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);
            if (parseStream.CanSeek) {
                parseStream.Position = 0;
            }

            PdfLogicalDocument document = LoadDocument(parseStream, effectivePdfOptions);
            foreach (var chunk in ReadPdf(document, source, effectiveReaderOptions, effectivePdfOptions, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into normalized Reader chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadPdf(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        foreach (var chunk in ReadPdf(document, source, effectiveReaderOptions, effectivePdfOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(PdfLogicalDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderPdfOptions pdfOptions, CancellationToken cancellationToken) {
        int maxChars = readerOptions.MaxChars > 0 ? readerOptions.MaxChars : 8_000;
        var markdownOptions = ReaderPdfOptions.CloneMarkdownOptions(pdfOptions.MarkdownOptions) ?? ReaderPdfOptions.CreateOfficeIMOProfile().MarkdownOptions!;
        markdownOptions.IncludePageSeparators = false;

        if (!pdfOptions.ChunkByPage) {
            string markdown = document.ToMarkdown(markdownOptions);
            foreach (var chunk in BuildChunksFromText(markdown, source, readerOptions, page: null, sourceBlockIndex: 0, blockKind: "document", blockAnchor: "document", tables: BuildTables(document.Tables, readerOptions), idPrefix: "pdf-document", maxChars: maxChars, cancellationToken: cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        int emittedIndex = 0;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();

            PdfLogicalPage page = document.Pages[pageIndex];
            string markdown = page.ToMarkdown(markdownOptions);
            var pageTables = BuildTables(page.Tables, readerOptions);
            foreach (var chunk in BuildChunksFromText(markdown, source, readerOptions, page.PageNumber, pageIndex, "page", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture), pageTables, "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture), maxChars, cancellationToken)) {
                chunk.Location.BlockIndex = emittedIndex++;
                yield return chunk;
            }
        }

        if (emittedIndex == 0) {
            const string warning = "PDF content produced no readable logical pages.";
            yield return EnrichChunk(new ReaderChunk {
                Id = "pdf-warning-0000",
                Kind = ReaderInputKind.Pdf,
                Location = new ReaderLocation {
                    Path = source.Path,
                    BlockIndex = 0,
                    SourceBlockKind = "warning"
                },
                Text = warning,
                Warnings = new[] { warning }
            }, source, readerOptions.ComputeHashes);
        }
    }

    private static IEnumerable<ReaderChunk> BuildChunksFromText(string markdown, SourceMetadata source, ReaderOptions readerOptions, int? page, int sourceBlockIndex, string blockKind, string blockAnchor, IReadOnlyList<ReaderTable>? tables, string idPrefix, int maxChars, CancellationToken cancellationToken) {
        var parts = SplitText(markdown, maxChars);
        if (parts.Count == 0) {
            string warning = page.HasValue
                ? "PDF page " + page.Value.ToString(CultureInfo.InvariantCulture) + " produced no readable text."
                : "PDF content produced no readable text.";
            parts = new[] { new TextPart(warning, new[] { warning }) };
        }

        for (int i = 0; i < parts.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            TextPart part = parts[i];
            yield return EnrichChunk(new ReaderChunk {
                Id = parts.Count == 1
                    ? idPrefix
                    : idPrefix + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Pdf,
                Location = new ReaderLocation {
                    Path = source.Path,
                    SourceBlockIndex = sourceBlockIndex,
                    SourceBlockKind = blockKind,
                    BlockAnchor = parts.Count == 1 ? blockAnchor : blockAnchor + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                    Page = page
                },
                Text = part.Text,
                Markdown = part.Text,
                Tables = i == 0 ? tables : null,
                Warnings = part.Warnings
            }, source, readerOptions.ComputeHashes);
        }
    }

    private static IReadOnlyList<TextPart> SplitText(string text, int maxChars) {
        if (string.IsNullOrWhiteSpace(text)) return Array.Empty<TextPart>();
        if (maxChars <= 0 || text.Length <= maxChars) return new[] { new TextPart(text.Trim(), null) };

        var parts = new List<TextPart>();
        int index = 0;
        while (index < text.Length) {
            int remaining = text.Length - index;
            int take = Math.Min(maxChars, remaining);
            if (take < remaining) {
                int splitAt = FindSplit(text, index, take);
                if (splitAt > index) {
                    take = splitAt - index;
                }
            }

            string segment = text.Substring(index, take).Trim();
            if (segment.Length > 0) {
                parts.Add(new TextPart(segment, new[] { "PDF content was split due to MaxChars." }));
            }

            index += take;
            while (index < text.Length && char.IsWhiteSpace(text[index])) {
                index++;
            }
        }

        return parts;
    }

    private static int FindSplit(string text, int index, int take) {
        int end = Math.Min(text.Length, index + take);
        for (int i = end - 1; i > index; i--) {
            char ch = text[i];
            if (ch == '\n' || ch == '\r' || char.IsWhiteSpace(ch)) {
                return i;
            }
        }

        return end;
    }

    private static IReadOnlyList<ReaderTable>? BuildTables(IReadOnlyList<PdfLogicalTable> tables, ReaderOptions readerOptions) {
        if (tables.Count == 0) return null;

        var result = new List<ReaderTable>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            PdfLogicalTable table = tables[i];
            int totalRows = table.Rows.Count;
            int maxRows = readerOptions.MaxTableRows > 0 ? readerOptions.MaxTableRows : totalRows;
            var rows = table.Rows.Take(maxRows).Select(static row => (IReadOnlyList<string>)row.ToArray()).ToArray();
            int columnCount = 0;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                columnCount = Math.Max(columnCount, table.Rows[rowIndex].Count);
            }

            result.Add(new ReaderTable {
                Kind = table.DetectionKind,
                Columns = Enumerable.Range(1, columnCount)
                    .Select(static column => "Column " + column.ToString(CultureInfo.InvariantCulture))
                    .ToArray(),
                Rows = rows,
                TotalRowCount = totalRows,
                Truncated = rows.Length < totalRows
            });
        }

        return result;
    }

    private static PdfLogicalDocument LoadDocument(string path, ReaderPdfOptions options) {
        var ranges = options.PageRanges?.ToArray();
        return ranges is { Length: > 0 }
            ? PdfLogicalDocument.LoadPageRanges(path, options.LayoutOptions, ranges)
            : PdfLogicalDocument.Load(path, options.LayoutOptions);
    }

    private static PdfLogicalDocument LoadDocument(Stream stream, ReaderPdfOptions options) {
        var ranges = options.PageRanges?.ToArray();
        return ranges is { Length: > 0 }
            ? PdfLogicalDocument.LoadPageRanges(stream, options.LayoutOptions, ranges)
            : PdfLogicalDocument.Load(stream, options.LayoutOptions);
    }

    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceMetadata source, bool computeHashes) {
        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
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
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
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

    private static string NormalizeLogicalSourceName(string? sourceName, string fallback) {
        if (sourceName is null) return fallback;
        string trimmed = sourceName.Trim();
        return trimmed.Length == 0 ? fallback : trimmed;
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

    private sealed class TextPart {
        public TextPart(string text, IReadOnlyList<string>? warnings) {
            Text = text;
            Warnings = warnings;
        }

        public string Text { get; }
        public IReadOnlyList<string>? Warnings { get; }
    }
}
