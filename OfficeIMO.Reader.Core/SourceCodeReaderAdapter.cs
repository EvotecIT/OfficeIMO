using System.Globalization;

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
        RejectBinary(input.Bytes, sourceName);
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
        if (ReaderSourceCodeLanguages.TryGetByExtension(DocumentReaderEngine.NormalizeExtension(ReaderLogicalPath.GetExtension(sourceName)),
                out ReaderSourceCodeLanguages.Language byExtension)) {
            language = byExtension;
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
            int lineCount = CountLines(part);
            var chunk = new ReaderChunk {
                Id = "code:" + fileName + ":" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Text,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = index,
                    StartLine = line,
                    EndLine = line + Math.Max(0, lineCount - 1),
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
            line += lineCount;
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

    /// <summary>Refuses payloads that carry NUL bytes outside UTF-16 or UTF-32 text.</summary>
    private static void RejectBinary(byte[] bytes, string sourceName) {
        if (HasWideUnicodeBom(bytes)) return;
        if (Array.IndexOf(bytes, (byte)0) >= 0) throw BinaryContent(sourceName);
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
        return fence + (language ?? string.Empty) + "\n" + body + fence;
    }
}
