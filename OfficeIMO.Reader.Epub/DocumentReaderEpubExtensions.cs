using System.Runtime.InteropServices;
using OfficeIMO.Epub;

namespace OfficeIMO.Reader.Epub;

/// <summary>
/// EPUB ingestion adapter for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderEpubExtensions {
    /// <summary>
    /// Reads EPUB content from a file path and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadEpub(string epubPath, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubPath == null) throw new ArgumentNullException(nameof(epubPath));
        if (epubPath.Length == 0) throw new ArgumentException("EPUB path cannot be empty.", nameof(epubPath));
        if (!File.Exists(epubPath)) throw new FileNotFoundException($"EPUB file '{epubPath}' doesn't exist.", epubPath);

        var options = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(epubPath, options.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(epubPath, options.ComputeHashes);
        var document = EpubReader.Read(epubPath, epubOptions);
        return ReadEpubDocument(document, source, options, cancellationToken);
    }

    /// <summary>
    /// Reads EPUB content from a stream and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadEpub(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new ArgumentException("EPUB stream must be readable.", nameof(epubStream));

        var options = readerOptions ?? new ReaderOptions();
        var logicalSourceName = "document.epub";
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                logicalSourceName = trimmedSourceName;
            }
        }
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(epubStream, options.MaxInputBytes, cancellationToken, out var ownsParseStream);
        var source = BuildSourceMetadataFromStream(parseStream, logicalSourceName, options.ComputeHashes);
        EpubDocument document;
        try {
            document = EpubReader.Read(parseStream, epubOptions);
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }

        return ReadEpubDocument(document, source, options, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadEpubDocument(EpubDocument document, SourceMetadata source, ReaderOptions options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (options == null) throw new ArgumentNullException(nameof(options));

        int maxChars = options.MaxChars > 0 ? options.MaxChars : 8_000;
        var sourcePath = source.Path;
        var fileName = Path.GetFileName(sourcePath);

        int warningIndex = 0;
        foreach (var warning in document.Warnings) {
            cancellationToken.ThrowIfCancellationRequested();
            var warningChunk = EnrichChunk(BuildWarningChunk(sourcePath, warning, warningIndex++), source, options.ComputeHashes);
            ApplyVirtualSourceMetadata(warningChunk, warningChunk.Location.Path, options.ComputeHashes);
            yield return warningChunk;
        }

        int blockIndex = warningIndex;
        foreach (var chapter in document.Chapters) {
            cancellationToken.ThrowIfCancellationRequested();

            int chunkPart = 0;
            foreach (var piece in SplitText(chapter.Text, maxChars)) {
                cancellationToken.ThrowIfCancellationRequested();

                var chunkId = BuildId(fileName, chapter.Order, chunkPart);
                var wasSplit = chapter.Text.Length > maxChars;

                var chunk = EnrichChunk(new ReaderChunk {
                    Id = chunkId,
                    Kind = ReaderInputKind.Epub,
                    Location = new ReaderLocation {
                        Path = BuildVirtualPath(sourcePath, chapter.Path),
                        BlockIndex = blockIndex,
                        SourceBlockIndex = chapter.Order > 0 ? chapter.Order - 1 : null,
                        HeadingPath = chapter.Title
                    },
                    Text = piece,
                    Markdown = BuildMarkdown(chapter.Title, piece),
                    Warnings = wasSplit ? new[] { "Chapter content was split due to MaxChars." } : null
                }, source, options.ComputeHashes);
                ApplyVirtualSourceMetadata(chunk, chunk.Location.Path, options.ComputeHashes);
                yield return chunk;

                blockIndex++;
                chunkPart++;
            }
        }
    }

    private static string BuildMarkdown(string? title, string text) {
        var heading = title;
        if (heading == null) return text;

        var headingTrimmed = heading.Trim();
        if (headingTrimmed.Length == 0) return text;

        return "## " + headingTrimmed + Environment.NewLine + Environment.NewLine + text;
    }

    private static IEnumerable<string> SplitText(string text, int maxChars) {
        if (string.IsNullOrWhiteSpace(text)) yield break;
        if (text.Length <= maxChars) {
            yield return text;
            yield break;
        }

        int index = 0;
        while (index < text.Length) {
            int remaining = text.Length - index;
            int take = Math.Min(maxChars, remaining);
            int end = index + take;

            if (end < text.Length) {
                int split = text.LastIndexOf(' ', end - 1, take);
                if (split > index + 128) {
                    end = split;
                }
            }

            var piece = text.Substring(index, end - index).Trim();
            if (piece.Length > 0) yield return piece;

            index = end;
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
        }
    }

    private static string BuildId(string fileName, int chapterOrder, int chunkPart) {
        var safeFile = Sanitize(fileName);
        return string.Concat(
            "epub-",
            safeFile,
            "-ch",
            chapterOrder.ToString("D4", CultureInfo.InvariantCulture),
            "-",
            chunkPart.ToString("D4", CultureInfo.InvariantCulture));
    }

    private static ReaderChunk BuildWarningChunk(string epubPath, string warning, int warningIndex) {
        var warningPath = ResolveWarningPath(epubPath, warning);

        return new ReaderChunk {
            Id = $"epub-warning-{warningIndex.ToString("D4", CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Epub,
            Location = new ReaderLocation {
                Path = warningPath,
                BlockIndex = warningIndex
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static string ResolveWarningPath(string epubPath, string warning) {
        if (string.IsNullOrWhiteSpace(warning)) return epubPath;

        if (warning.Contains("container.xml", StringComparison.OrdinalIgnoreCase)) {
            return BuildVirtualPath(epubPath, "META-INF/container.xml");
        }

        var quotedValue = TryExtractQuotedValue(warning);
        if (IsArchiveRelativePath(quotedValue)) {
            return BuildVirtualPath(epubPath, quotedValue!);
        }

        return epubPath;
    }

    private static string? TryExtractQuotedValue(string text) {
        if (string.IsNullOrEmpty(text)) return null;

        var start = text.IndexOf('\'');
        if (start < 0 || start + 1 >= text.Length) return null;

        var end = text.IndexOf('\'', start + 1);
        if (end <= start + 1) return null;

        return text.Substring(start + 1, end - start - 1).Trim();
    }

    private static bool IsArchiveRelativePath(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (Path.IsPathRooted(value)) return false;

        var normalized = value!.Replace('\\', '/');
        return normalized.IndexOf('/') >= 0 || Path.GetExtension(normalized).Length > 0;
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

    private static void ApplyVirtualSourceMetadata(ReaderChunk chunk, string? virtualPath, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (string.IsNullOrWhiteSpace(virtualPath)) return;

        var effectiveVirtualPath = virtualPath!;
        chunk.Location.Path = effectiveVirtualPath;
        chunk.SourceId = BuildSourceId(effectiveVirtualPath);

        if (computeHashes) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }
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

    private static SourceMetadata BuildSourceMetadataFromPath(string epubPath, bool computeHash) {
        var normalizedPath = NormalizePathForId(epubPath);
        var sourceId = BuildSourceId(normalizedPath);

        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fileInfo = new FileInfo(epubPath);
            if (fileInfo.Exists) {
                lastWriteUtc = fileInfo.LastWriteTimeUtc;
                lengthBytes = fileInfo.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeFileSha256(epubPath);
        }

        return new SourceMetadata {
            Path = normalizedPath,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static SourceMetadata BuildSourceMetadataFromStream(Stream stream, string sourceName, bool computeHash) {
        var logicalName = string.IsNullOrWhiteSpace(sourceName) ? "document.epub" : sourceName.Trim();
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
        var normalized = NormalizeSourceKeyForId(sourceKey);

        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string NormalizeSourceKeyForId(string? sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar != '\\') {
            return normalized;
        }

        var virtualPathSeparatorIndex = normalized.IndexOf("::", StringComparison.Ordinal);
        if (virtualPathSeparatorIndex < 0) {
            return normalized.ToLowerInvariant();
        }

        var archivePath = normalized.Substring(0, virtualPathSeparatorIndex).ToLowerInvariant();
        var archiveEntryPath = normalized.Substring(virtualPathSeparatorIndex);
        return archivePath + archiveEntryPath;
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        string fullPath;
        try {
            fullPath = Path.GetFullPath(path);
        } catch {
            fullPath = path;
        }

        fullPath = ResolveExistingFullPath(fullPath);
        return fullPath.Replace('\\', '/');
    }

    private static string ResolveExistingFullPath(string fullPath) {
        if (Path.DirectorySeparatorChar == '\\') {
            return fullPath;
        }

        IntPtr resolvedPathPointer = IntPtr.Zero;
        try {
            resolvedPathPointer = UnixRealPath(fullPath, IntPtr.Zero);
            if (resolvedPathPointer == IntPtr.Zero) {
                return fullPath;
            }

            var resolvedPath = Marshal.PtrToStringAnsi(resolvedPathPointer);
            return string.IsNullOrWhiteSpace(resolvedPath) ? fullPath : resolvedPath;
        } catch {
            return fullPath;
        } finally {
            if (resolvedPathPointer != IntPtr.Zero) {
                UnixFree(resolvedPathPointer);
            }
        }
    }

    [DllImport("libc", EntryPoint = "realpath", CharSet = CharSet.Ansi)]
    private static extern IntPtr UnixRealPath(string path, IntPtr buffer);

    [DllImport("libc", EntryPoint = "free")]
    private static extern void UnixFree(IntPtr pointer);

    private static string BuildVirtualPath(string epubPath, string chapterPath) {
        if (string.IsNullOrWhiteSpace(chapterPath)) return epubPath;
        return epubPath + "::" + chapterPath.Replace('\\', '/');
    }

    private static string Sanitize(string value) {
        if (string.IsNullOrWhiteSpace(value)) return "document";

        var sb = new StringBuilder(value.Length);
        foreach (var ch in value) {
            if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) {
                sb.Append(char.ToLowerInvariant(ch));
            } else {
                sb.Append('-');
            }
        }

        return sb.ToString().Trim('-');
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
