using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Reader.Html;

/// <summary>
/// HTML ingestion adapter for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderHtmlExtensions {
    /// <summary>
    /// Reads an HTML file and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadHtmlFile(string htmlPath, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (htmlPath == null) throw new ArgumentNullException(nameof(htmlPath));
        if (htmlPath.Length == 0) throw new ArgumentException("HTML path cannot be empty.", nameof(htmlPath));
        if (!File.Exists(htmlPath)) throw new FileNotFoundException($"HTML file '{htmlPath}' doesn't exist.", htmlPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(htmlPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(htmlPath, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(htmlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadHtml(fs, source, effectiveReaderOptions, htmlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads an HTML stream and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadHtml(Stream htmlStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        if (!htmlStream.CanRead) throw new ArgumentException("HTML stream must be readable.", nameof(htmlStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.html" : sourceName.Trim();
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };
        foreach (var chunk in ReadHtml(htmlStream, source, effectiveReaderOptions, htmlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadHtml(Stream htmlStream, SourceMetadata source, ReaderOptions effectiveReaderOptions, ReaderHtmlOptions? htmlOptions, CancellationToken cancellationToken) {
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(htmlStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out var ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);
            var html = ReadAllText(parseStream, cancellationToken);
            foreach (var chunk in ReadHtmlString(html, source, effectiveReaderOptions, htmlOptions, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Reads an HTML string and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadHtmlString(string html, string sourceName = "document.html", ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effective = readerOptions ?? new ReaderOptions();
        var trimmedSourceName = sourceName.Trim();
        var logicalSourceName = trimmedSourceName.Length == 0 ? "document.html" : trimmedSourceName;
        var source = BuildSourceMetadataFromHtmlString(logicalSourceName, html, effective.ComputeHashes);

        foreach (var chunk in ReadHtmlString(html, source, effective, htmlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadHtmlString(string html, SourceMetadata source, ReaderOptions effective, ReaderHtmlOptions? htmlOptions, CancellationToken cancellationToken) {
        var effectiveHtmlOptions = ReaderHtmlOptionsCloner.CloneOrDefault(htmlOptions);
        int maxChars = effective.MaxChars > 0 ? effective.MaxChars : 8_000;
        var logicalSourceName = source.Path;

        string markdown = html.ToMarkdown(effectiveHtmlOptions.HtmlToMarkdownOptions);

        if (string.IsNullOrWhiteSpace(markdown)) {
            yield return EnrichChunk(BuildWarningChunk(logicalSourceName, "html-warning-0000", "HTML content produced no markdown text."), source, effective.ComputeHashes);
            yield break;
        }

        var parts = SplitMarkdown(
            markdown,
            maxChars,
            chunkByHeadings: effective.MarkdownChunkByHeadings,
            cancellationToken);

        int chunkIndex = 0;
        foreach (var part in parts) {
            cancellationToken.ThrowIfCancellationRequested();

            yield return EnrichChunk(new ReaderChunk {
                Id = string.Concat("html-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
                Kind = ReaderInputKind.Html,
                Location = new ReaderLocation {
                    Path = logicalSourceName,
                    BlockIndex = chunkIndex,
                    StartLine = part.StartLine,
                    HeadingPath = part.HeadingPath
                },
                Text = part.Text,
                Markdown = part.Text,
                Warnings = part.Warnings
            }, source, effective.ComputeHashes);

            chunkIndex++;
        }
    }

    private static IReadOnlyList<MarkdownPart> SplitMarkdown(string markdown, int maxChars, bool chunkByHeadings, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(markdown)) return Array.Empty<MarkdownPart>();

        var normalized = markdown
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        var lines = normalized.Split('\n');

        var parts = new List<MarkdownPart>(capacity: Math.Max(1, lines.Length / 16));
        var headingStack = new List<(int Level, string Text)>();
        var currentText = new StringBuilder(Math.Min(maxChars, 4 * 1024));
        var currentWarnings = new List<string>(2);
        int currentStartLine = 1;
        string? currentHeadingPath = null;

        void FlushCurrent() {
            if (currentText.Length == 0) return;

            var text = currentText.ToString().TrimEnd();
            if (text.Length > 0) {
                parts.Add(new MarkdownPart(
                    text,
                    currentStartLine,
                    currentHeadingPath,
                    currentWarnings.Count == 0 ? null : currentWarnings.ToArray()));
            }

            currentText.Clear();
            currentWarnings.Clear();
        }

        for (int i = 0; i < lines.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();

            var line = lines[i];
            int lineNo = i + 1;

            int headingLevel = 0;
            string headingText = string.Empty;
            bool isHeading = false;
            if (chunkByHeadings) {
                isHeading = TryParseAtxHeading(line, out headingLevel, out headingText);
            }

            if (isHeading && currentText.Length > 0) {
                FlushCurrent();
            }

            if (isHeading) {
                UpdateHeadingStack(headingStack, headingLevel, headingText);
            }

            if (currentText.Length == 0) {
                currentStartLine = lineNo;
                currentHeadingPath = chunkByHeadings ? BuildHeadingPath(headingStack) : null;
            }

            if (line.Length > maxChars) {
                if (currentText.Length > 0) {
                    AddSplitWarning(currentWarnings);
                    FlushCurrent();
                }

                int segmentIndex = 0;
                while (segmentIndex < line.Length) {
                    if (currentText.Length == 0) {
                        currentStartLine = lineNo;
                        currentHeadingPath = chunkByHeadings ? BuildHeadingPath(headingStack) : null;
                    }

                    int take = Math.Min(maxChars, line.Length - segmentIndex);
                    currentText.Append(line, segmentIndex, take);
                    segmentIndex += take;
                    AddSplitWarning(currentWarnings);

                    if (segmentIndex < line.Length) {
                        FlushCurrent();
                    }
                }

                continue;
            }

            if (WouldExceed(maxChars, currentText, line)) {
                AddSplitWarning(currentWarnings);
                FlushCurrent();
                currentStartLine = lineNo;
                currentHeadingPath = chunkByHeadings ? BuildHeadingPath(headingStack) : null;
            }

            AppendLine(currentText, line);
        }

        FlushCurrent();
        return parts;
    }

    private static string ReadAllText(Stream stream, CancellationToken cancellationToken) {
        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);

        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            var read = reader.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            sb.Append(buffer, 0, read);
        }

        return sb.ToString();
    }

    private static ReaderChunk BuildWarningChunk(string sourceName, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Html,
            Location = new ReaderLocation {
                Path = sourceName,
                BlockIndex = 0
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static bool TryParseAtxHeading(string line, out int level, out string text) {
        level = 0;
        text = string.Empty;
        if (line == null) return false;

        int i = 0;
        while (i < line.Length && line[i] == '#') i++;
        if (i < 1 || i > 6) return false;
        if (i >= line.Length) return false;
        if (line[i] != ' ' && line[i] != '\t') return false;

        level = i;
        text = line.Substring(i).Trim();
        if (text.Length == 0) text = "Heading " + level.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    private static void UpdateHeadingStack(List<(int Level, string Text)> stack, int level, string text) {
        if (level < 1) return;
        if (string.IsNullOrWhiteSpace(text)) text = "Heading " + level.ToString(CultureInfo.InvariantCulture);

        for (int i = stack.Count - 1; i >= 0; i--) {
            if (stack[i].Level >= level) stack.RemoveAt(i);
        }
        stack.Add((level, CollapseWhitespace(text)));
    }

    private static string? BuildHeadingPath(List<(int Level, string Text)> stack) {
        if (stack.Count == 0) return null;

        var sb = new StringBuilder();
        for (int i = 0; i < stack.Count; i++) {
            if (i > 0) sb.Append(" > ");
            sb.Append(stack[i].Text);
        }

        var value = sb.ToString().Trim();
        return value.Length == 0 ? null : value;
    }

    private static bool WouldExceed(int maxChars, StringBuilder current, string nextLine) {
        int nextLength = nextLine?.Length ?? 0;
        int extra = (current.Length == 0 ? 0 : 1) + nextLength;
        return current.Length > 0 && (current.Length + extra) > maxChars;
    }

    private static void AppendLine(StringBuilder builder, string line) {
        if (builder.Length > 0) builder.AppendLine();
        builder.Append(line ?? string.Empty);
    }

    private static void AddSplitWarning(List<string> warnings) {
        const string splitWarning = "HTML content was split due to MaxChars.";
        for (int i = 0; i < warnings.Count; i++) {
            if (string.Equals(warnings[i], splitWarning, StringComparison.Ordinal)) {
                return;
            }
        }

        warnings.Add(splitWarning);
    }

    private static string CollapseWhitespace(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var sb = new StringBuilder(value.Length);
        bool previousWhitespace = false;
        for (int i = 0; i < value.Length; i++) {
            var ch = value[i];
            bool isWhitespace = char.IsWhiteSpace(ch);
            if (isWhitespace) {
                if (!previousWhitespace) sb.Append(' ');
                previousWhitespace = true;
            } else {
                sb.Append(ch);
                previousWhitespace = false;
            }
        }

        return sb.ToString().Trim();
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

    private static SourceMetadata BuildSourceMetadataFromHtmlString(string sourcePath, string html, bool computeHash) {
        return new SourceMetadata {
            Path = sourcePath,
            SourceId = BuildSourceId(sourcePath),
            SourceHash = computeHash ? ComputeSha256Hex(html) : null,
            LengthBytes = Encoding.UTF8.GetByteCount(html)
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

    private sealed class MarkdownPart {
        public MarkdownPart(string text, int startLine, string? headingPath, IReadOnlyList<string>? warnings) {
            Text = text;
            StartLine = startLine;
            HeadingPath = headingPath;
            Warnings = warnings;
        }

        public string Text { get; }
        public int StartLine { get; }
        public string? HeadingPath { get; }
        public IReadOnlyList<string>? Warnings { get; }
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
