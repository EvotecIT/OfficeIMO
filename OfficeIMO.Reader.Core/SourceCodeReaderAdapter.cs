using System.Globalization;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>Projects source-code files as verbatim text, fenced Markdown and a language metadata entry.</summary>
internal static class SourceCodeReaderAdapter {
    private const string MetadataCategory = "source-code.document";

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions options, CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(path, options, cancellationToken);
        return BuildDocument(input, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(stream, sourceName, options, cancellationToken);
        return BuildDocument(input, options, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocument(ReaderAdapterInputSnapshot input, ReaderOptions options,
        CancellationToken cancellationToken) {
        string sourceName = input.Source.Path ?? "memory";
        if (LooksBinary(input.Bytes, input.Bytes.Length, options)) throw BinaryContent(sourceName);
        var decoding = new ReaderTextDecoding();
        string text;
        using (var stream = new MemoryStream(input.Bytes, writable: false))
        using (TextReader reader = decoding.Open(stream, options, cancellationToken)) {
            text = reader.ReadToEnd();
        }
        if (text.IndexOf('\0') >= 0) throw BinaryContent(sourceName);
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');

        string detectedBy = "extension";
        ReaderSourceCodeLanguages.Language? language = null;
        string extension = DocumentReaderEngine.NormalizeExtension(ReaderLogicalPath.GetExtension(sourceName));
        if (ReaderSourceCodeLanguages.TryGetByExtension(extension, out ReaderSourceCodeLanguages.Language byExtension)) {
            language = byExtension;
        } else if (extension.Length == 0 &&
                   ReaderSourceCodeLanguages.TryGetExtensionFromFileName(ReaderLogicalPath.GetFileName(sourceName), out string namedExtension) &&
                   ReaderSourceCodeLanguages.TryGetByExtension(namedExtension, out ReaderSourceCodeLanguages.Language byFileName)) {
            language = byFileName;
            detectedBy = "file-name";
        } else if (ReaderSourceCodeLanguages.TryGetExtensionFromShebang(text, out string scriptExtension) &&
                   ReaderSourceCodeLanguages.TryGetByExtension(scriptExtension, out ReaderSourceCodeLanguages.Language byShebang)) {
            language = byShebang;
            detectedBy = "shebang";
        }

        bool emitMarkdown = options.Projection != ReaderProjection.TextAndMetadata;
        string fileName = ReaderLogicalPath.GetFileName(sourceName);
        var chunks = new List<ReaderChunk>();
        int index = 0;
        int line = 1;
        foreach (string part in SplitOnLines(text, Math.Max(256, options.MaxChars))) {
            cancellationToken.ThrowIfCancellationRequested();
            // A part that ends inside a long line leaves the next part on that same line.
            int newLines = CountNewLines(part);
            int endLine = part.Length == 0 ? line : line + newLines - (part[part.Length - 1] == '\n' ? 1 : 0);
            var chunk = new ReaderChunk {
                Id = "code:" + fileName + ":" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Text,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = index,
                    StartLine = line,
                    EndLine = endLine,
                    SourceBlockKind = "source-code"
                },
                Text = part,
                Markdown = emitMarkdown ? Fence(part, language?.Id) : null,
                ContinuesPreviousChunk = index > 0
            };
            if (index == 0 && decoding.InvalidSequences > 0) {
                chunk.Warnings = new[] { "Source input contained invalid byte sequences; they were replaced with U+FFFD." };
            }
            DocumentReaderEngine.ApplyAdapterSource(chunk, input, options.ComputeHashes);
            chunks.Add(chunk);
            line += newLines;
            index++;
        }

        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(chunks, ReaderInputKind.Text, input.Source,
            new[] { OfficeDocumentReaderBuilderSourceCodeExtensions.HandlerId });
        result.Source.Title = fileName;
        result.Metadata = result.Metadata.Concat(BuildMetadata(language, detectedBy, CountLines(text))).ToArray();
        return result;
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(ReaderSourceCodeLanguages.Language? language,
        string detectedBy, int lineCount) {
        if (language != null) {
            yield return new OfficeDocumentMetadataEntry {
                Id = "source-code-language",
                Category = MetadataCategory,
                Name = "Language",
                Value = language.Id,
                ValueType = "string",
                Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["displayName"] = language.Name,
                    ["detectedBy"] = detectedBy
                }
            };
        }
        yield return new OfficeDocumentMetadataEntry {
            Id = "source-code-line-count",
            Category = MetadataCategory,
            Name = "LineCount",
            Value = lineCount.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        };
    }

    /// <summary>
    /// Validates a claimed source extension: the prefix must be text the adapter accepts. The caller restores the
    /// stream position.
    /// </summary>
    internal static bool Probe(Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        var prefix = new byte[ProbeBytes];
        int count = 0;
        while (count < prefix.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(prefix, count, prefix.Length - count);
            if (read == 0) break;
            count += read;
        }
        return !LooksBinary(prefix, count, options);
    }

    private const int ProbeBytes = 8192;

    /// <summary>
    /// NUL bytes mark binary data unless the text is UTF-16 or UTF-32, announced by a BOM or by the configured
    /// <see cref="ReaderOptions.TextEncoding"/>; NUL characters remaining after decoding are refused separately.
    /// </summary>
    private static bool LooksBinary(byte[] bytes, int count, ReaderOptions options) {
        if (HasWideUnicodeBom(bytes) || options.TextEncoding is UnicodeEncoding or UTF32Encoding) return false;
        return Array.IndexOf(bytes, (byte)0, 0, count) >= 0;
    }

    private static bool HasWideUnicodeBom(byte[] bytes) =>
        bytes.Length >= 2 && ((bytes[0] == 0xFF && bytes[1] == 0xFE) || (bytes[0] == 0xFE && bytes[1] == 0xFF)) ||
        bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xFE && bytes[3] == 0xFF;

    private static InvalidDataException BinaryContent(string sourceName) =>
        new("'" + ReaderLogicalPath.GetFileName(sourceName) + "' contains binary data and was not read as source code.");

    /// <summary>Splits on line boundaries so a chunk never ends inside a line unless one line exceeds the limit.</summary>
    private static IEnumerable<string> SplitOnLines(string text, int limit) {
        if (text.Length == 0) {
            yield return string.Empty;
            yield break;
        }
        int offset = 0;
        while (offset < text.Length) {
            int remaining = text.Length - offset;
            if (remaining <= limit) {
                yield return text.Substring(offset);
                yield break;
            }
            int lastBreak = text.LastIndexOf('\n', offset + limit - 1, limit);
            int length = lastBreak >= offset ? lastBreak - offset + 1 : limit;
            if (lastBreak < offset && char.IsHighSurrogate(text[offset + length - 1])) length--;
            yield return text.Substring(offset, length);
            offset += length;
        }
    }

    private static int CountNewLines(string text) {
        int count = 0;
        foreach (char character in text) {
            if (character == '\n') count++;
        }
        return count;
    }

    private static int CountLines(string text) {
        if (text.Length == 0) return 0;
        int count = 1;
        for (int index = 0; index < text.Length - 1; index++) {
            if (text[index] == '\n') count++;
        }
        return count;
    }

    private static string Fence(string code, string? language) {
        int longestRun = 0;
        int run = 0;
        foreach (char character in code) {
            run = character == '`' ? run + 1 : 0;
            if (run > longestRun) longestRun = run;
        }
        string fence = new('`', Math.Max(3, longestRun + 1));
        string body = code.EndsWith("\n", StringComparison.Ordinal) ? code : code + "\n";
        // The closing fence ends its line, so continuation chunks joined into document Markdown open a new fence.
        return fence + (language ?? string.Empty) + "\n" + body + fence + "\n";
    }
}
