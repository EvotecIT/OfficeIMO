using System.Globalization;
using System.Threading;

namespace OfficeIMO.Reader;

internal static class TextReaderAdapter {
    internal static IReadOnlyList<ReaderChunk> Read(string path, ReaderInputKind kind, ReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Read(stream, path, kind, options, cancellationToken);
    }

    internal static IReadOnlyList<ReaderChunk> Read(Stream stream, string? sourceName, ReaderInputKind kind, ReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) stream.Position = 0;
        using var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, leaveOpen: true);
        string text = reader.ReadToEnd();
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : sourceName!;
        return Chunk(text, logicalName, kind, options.MaxChars, cancellationToken).ToArray();
    }

    private static IEnumerable<ReaderChunk> Chunk(string text, string sourceName, ReaderInputKind kind, int maxChars, CancellationToken cancellationToken) {
        int limit = Math.Max(256, maxChars);
        string normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Length == 0) {
            yield return CreateChunk(string.Empty, sourceName, kind, 0, 1);
            yield break;
        }

        int index = 0;
        int line = 1;
        for (int offset = 0; offset < normalized.Length; offset += limit) {
            cancellationToken.ThrowIfCancellationRequested();
            string part = normalized.Substring(offset, Math.Min(limit, normalized.Length - offset));
            yield return CreateChunk(part, sourceName, kind, index++, line);
            line += part.Count(static character => character == '\n');
        }
    }

    private static ReaderChunk CreateChunk(string text, string sourceName, ReaderInputKind kind, int index, int startLine) =>
        new ReaderChunk {
            Id = $"{(kind == ReaderInputKind.Unknown ? "unknown" : "text")}:{ReaderLogicalPath.GetFileName(sourceName)}:{index.ToString("D4", CultureInfo.InvariantCulture)}",
            Kind = kind,
            Location = new ReaderLocation { Path = sourceName, BlockIndex = index, StartLine = startLine, SourceBlockKind = kind == ReaderInputKind.Unknown ? "unknown" : "text" },
            Text = text,
            Warnings = kind == ReaderInputKind.Unknown ? new[] { "Input was projected through the explicit unknown-payload fallback." } : null
        };
}
