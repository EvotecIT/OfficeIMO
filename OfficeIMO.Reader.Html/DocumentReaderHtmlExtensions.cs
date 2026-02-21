using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

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

        using var fs = new FileStream(htmlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadHtml(fs, htmlPath, readerOptions, htmlOptions, cancellationToken)) {
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
        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(htmlStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out var ownsParseStream);
        try {
            var html = ReadAllText(parseStream, cancellationToken);
            var logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.html" : sourceName!;
            foreach (var chunk in ReadHtmlString(html, logicalSourceName, effectiveReaderOptions, htmlOptions, cancellationToken)) {
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
        var effectiveHtmlOptions = ReaderHtmlOptionsCloner.CloneOrDefault(htmlOptions);
        int maxChars = effective.MaxChars > 0 ? effective.MaxChars : 8_000;
        var logicalSourceName = sourceName.Trim().Length == 0 ? "document.html" : sourceName;

        string markdown;
        using (var document = html.LoadFromHtml(effectiveHtmlOptions.HtmlToWordOptions)) {
            markdown = document.ToMarkdown(effectiveHtmlOptions.MarkdownOptions);
        }

        if (string.IsNullOrWhiteSpace(markdown)) {
            yield return BuildWarningChunk(logicalSourceName, "html-warning-0000", "HTML content produced no markdown text.");
            yield break;
        }

        var wasSplit = markdown.Length > maxChars;
        int chunkIndex = 0;
        foreach (var part in SplitText(markdown, maxChars)) {
            cancellationToken.ThrowIfCancellationRequested();

            yield return new ReaderChunk {
                Id = string.Concat("html-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
                Kind = ReaderInputKind.Unknown,
                Location = new ReaderLocation {
                    Path = logicalSourceName,
                    BlockIndex = chunkIndex
                },
                Text = part,
                Markdown = part,
                Warnings = wasSplit ? new[] { "HTML content was split due to MaxChars." } : null
            };

            chunkIndex++;
        }
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
                int split = text.LastIndexOf('\n', end - 1, take);
                if (split > index + 64) {
                    end = split;
                }
            }

            var part = text.Substring(index, end - index).Trim();
            if (part.Length > 0) yield return part;

            index = end;
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
        }
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
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation {
                Path = sourceName,
                BlockIndex = 0
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }
}
