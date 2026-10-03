using System.Globalization;
using System.Threading;

namespace OfficeIMO.Reader;

internal static class TextReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(string path, ReaderInputKind kind, ReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (ReaderChunk chunk in Read(stream, path, kind, options, cancellationToken)) yield return chunk;
    }

    internal static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName, ReaderInputKind kind, ReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) stream.Position = 0;
        var decoding = new ReaderTextDecoding();
        using var reader = decoding.Open(stream, options, cancellationToken);
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : sourceName!;
        foreach (ReaderChunk chunk in Chunk(reader, logicalName, kind, options.MaxChars, decoding, cancellationToken)) yield return chunk;
    }

    private static IEnumerable<ReaderChunk> Chunk(TextReader reader, string sourceName, ReaderInputKind kind, int maxChars,
        ReaderTextDecoding decoding, CancellationToken cancellationToken) {
        int limit = Math.Max(256, maxChars);
        var text = new StringBuilder(Math.Min(limit, 4096) + 1);
        int index = 0;
        int line = 1;
        bool afterCr = false;
        int observedInvalidSequences = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int value = reader.Read();
            if (value < 0) break;
            if (afterCr && value == '\n') { afterCr = false; continue; }
            afterCr = value == '\r';
            text.Append(afterCr ? '\n' : (char)value);
            if (text.Length < limit) continue;
            int length = char.IsHighSurrogate(text[text.Length - 1]) ? text.Length - 1 : text.Length;
            string part = text.ToString(0, length);
            text.Remove(0, length);
            ReaderChunk chunk = CreateChunk(part, sourceName, kind, index++, line);
            AddDecodingWarning(chunk, decoding, ref observedInvalidSequences);
            yield return chunk;
            line += part.Count(static character => character == '\n');
        }
        if (text.Length > 0 || index == 0) {
            ReaderChunk chunk = CreateChunk(text.ToString(), sourceName, kind, index, line);
            AddDecodingWarning(chunk, decoding, ref observedInvalidSequences);
            yield return chunk;
        }
    }

    private static void AddDecodingWarning(ReaderChunk chunk, ReaderTextDecoding decoding, ref int observed) {
        if (decoding.InvalidSequences <= observed) return;
        chunk.Warnings = (chunk.Warnings ?? Array.Empty<string>()).Concat(new[] {
            "Plain-text input contained invalid byte sequences; they were replaced with U+FFFD."
        }).ToArray();
        observed = decoding.InvalidSequences;
    }

    private static ReaderChunk CreateChunk(string text, string sourceName, ReaderInputKind kind, int index, int startLine) =>
        new ReaderChunk {
            Id = $"{(kind == ReaderInputKind.Unknown ? "unknown" : "text")}:{ReaderLogicalPath.GetFileName(sourceName)}:{index.ToString("D4", CultureInfo.InvariantCulture)}",
            Kind = kind,
            Location = new ReaderLocation { Path = sourceName, BlockIndex = index, StartLine = startLine, SourceBlockKind = kind == ReaderInputKind.Unknown ? "unknown" : "text" },
            Text = text,
            ContinuesPreviousChunk = index > 0,
            Warnings = kind == ReaderInputKind.Unknown ? new[] { "Input was projected through the explicit unknown-payload fallback." } : null
        };
}
