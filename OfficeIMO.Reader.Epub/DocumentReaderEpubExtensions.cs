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
        EnforceFileSize(epubPath, options.MaxInputBytes);
        var document = EpubReader.Read(epubPath, epubOptions);
        return ReadEpubDocument(document, epubPath, options, cancellationToken);
    }

    /// <summary>
    /// Reads EPUB content from a stream and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadEpub(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new ArgumentException("EPUB stream must be readable.", nameof(epubStream));

        var options = readerOptions ?? new ReaderOptions();
        var parseStream = EnsureSeekableReadStream(epubStream, options.MaxInputBytes, cancellationToken, out var ownsParseStream);
        EpubDocument document;
        try {
            document = EpubReader.Read(parseStream, epubOptions);
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }

        var logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.epub" : sourceName!;
        return ReadEpubDocument(document, logicalSourceName, options, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadEpubDocument(EpubDocument document, string sourcePath, ReaderOptions options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourcePath == null) throw new ArgumentNullException(nameof(sourcePath));
        if (options == null) throw new ArgumentNullException(nameof(options));

        int maxChars = options.MaxChars > 0 ? options.MaxChars : 8_000;
        var fileName = Path.GetFileName(sourcePath);

        int warningIndex = 0;
        foreach (var warning in document.Warnings) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return BuildWarningChunk(sourcePath, warning, warningIndex++);
        }

        int blockIndex = 0;
        foreach (var chapter in document.Chapters) {
            cancellationToken.ThrowIfCancellationRequested();

            int chunkPart = 0;
            foreach (var piece in SplitText(chapter.Text, maxChars)) {
                cancellationToken.ThrowIfCancellationRequested();

                var chunkId = BuildId(fileName, chapter.Order, chunkPart);
                var wasSplit = chapter.Text.Length > maxChars;

                yield return new ReaderChunk {
                    Id = chunkId,
                    Kind = ReaderInputKind.Unknown,
                    Location = new ReaderLocation {
                        Path = BuildVirtualPath(sourcePath, chapter.Path),
                        BlockIndex = blockIndex,
                        HeadingPath = chapter.Title
                    },
                    Text = piece,
                    Markdown = BuildMarkdown(chapter.Title, piece),
                    Warnings = wasSplit ? new[] { "Chapter content was split due to MaxChars." } : null
                };

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
        return new ReaderChunk {
            Id = $"epub-warning-{warningIndex.ToString("D4", CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation {
                Path = epubPath,
                BlockIndex = warningIndex
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

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

    private static Stream EnsureSeekableReadStream(Stream stream, long? maxInputBytes, CancellationToken cancellationToken, out bool ownsStream) {
        if (stream.CanSeek) {
            EnforceSeekableStreamSize(stream, maxInputBytes);
            ownsStream = false;
            return stream;
        }

        var buffer = new MemoryStream();
        try {
            var chunk = new byte[64 * 1024];
            long totalBytes = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                var read = stream.Read(chunk, 0, chunk.Length);
                if (read <= 0) break;
                buffer.Write(chunk, 0, read);

                totalBytes += read;
                if (maxInputBytes.HasValue && totalBytes > maxInputBytes.Value) {
                    throw new IOException(
                        $"Input exceeds MaxInputBytes ({totalBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}).");
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

    private static void EnforceFileSize(string path, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        var fi = new FileInfo(path);
        if (fi.Length > maxBytes.Value) {
            throw new IOException(
                $"Input exceeds MaxInputBytes ({fi.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}).");
        }
    }

    private static void EnforceSeekableStreamSize(Stream stream, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        if (!stream.CanSeek) return;
        try {
            if (stream.Length > maxBytes.Value) {
                throw new IOException(
                    $"Input exceeds MaxInputBytes ({stream.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}).");
            }
        } catch (NotSupportedException) {
            // ignore
        }
    }
}
