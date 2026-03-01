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

        var parts = SplitMarkdown(
            markdown,
            maxChars,
            chunkByHeadings: effective.MarkdownChunkByHeadings,
            cancellationToken);

        var wasSplit = parts.Count > 1;
        int chunkIndex = 0;
        foreach (var part in parts) {
            cancellationToken.ThrowIfCancellationRequested();

            IReadOnlyList<string>? warnings = null;
            if (wasSplit || (part.Warnings?.Count > 0)) {
                var warningList = new List<string>(4);
                if (wasSplit) {
                    warningList.Add("HTML content was split due to MaxChars.");
                }

                if (part.Warnings != null) {
                    foreach (var warning in part.Warnings) {
                        bool exists = false;
                        for (int i = 0; i < warningList.Count; i++) {
                            if (string.Equals(warningList[i], warning, StringComparison.Ordinal)) {
                                exists = true;
                                break;
                            }
                        }

                        if (!exists) {
                            warningList.Add(warning);
                        }
                    }
                }

                warnings = warningList.ToArray();
            }

            yield return new ReaderChunk {
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
                Warnings = warnings
            };

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

                    if (segmentIndex < line.Length) {
                        FlushCurrent();
                    }
                }

                continue;
            }

            if (WouldExceed(maxChars, currentText, line)) {
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
}
