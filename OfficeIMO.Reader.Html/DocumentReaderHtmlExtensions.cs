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
    public static IEnumerable<ReaderChunk> ReadHtmlFile(string htmlPath, ReaderOptions? readerOptions = null, CancellationToken cancellationToken = default) {
        if (htmlPath == null) throw new ArgumentNullException(nameof(htmlPath));
        if (htmlPath.Length == 0) throw new ArgumentException("HTML path cannot be empty.", nameof(htmlPath));
        if (!File.Exists(htmlPath)) throw new FileNotFoundException($"HTML file '{htmlPath}' doesn't exist.", htmlPath);

        var html = File.ReadAllText(htmlPath, Encoding.UTF8);
        foreach (var chunk in ReadHtmlString(html, htmlPath, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads an HTML string and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadHtmlString(string html, string sourceName = "document.html", ReaderOptions? readerOptions = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effective = readerOptions ?? new ReaderOptions();
        int maxChars = effective.MaxChars > 0 ? effective.MaxChars : 8_000;

        string markdown;
        using (var document = html.LoadFromHtml()) {
            markdown = document.ToMarkdown();
        }

        int chunkIndex = 0;
        foreach (var part in SplitText(markdown, maxChars)) {
            cancellationToken.ThrowIfCancellationRequested();

            yield return new ReaderChunk {
                Id = string.Concat("html-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
                Kind = ReaderInputKind.Unknown,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = chunkIndex
                },
                Text = part,
                Markdown = part,
                Warnings = markdown.Length > maxChars ? new[] { "HTML content was split due to MaxChars." } : null
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
}
