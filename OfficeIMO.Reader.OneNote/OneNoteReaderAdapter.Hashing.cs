using OfficeIMO.OneNote;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    private static string BuildChunkHashInput(ReaderChunk chunk) {
        return string.Join("|", chunk.Kind.ToString(), chunk.SourceId ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty, chunk.Text, chunk.Markdown ?? string.Empty);
    }

    private static string BuildSourceId(string value) {
        string normalized = Path.DirectorySeparatorChar == '\\' ? value.ToLowerInvariant() : value;
        return "src:" + ComputeHash(Encoding.UTF8.GetBytes(normalized));
    }

    private static string NormalizePath(string path) {
        try { return Path.GetFullPath(path).Replace('\\', '/'); }
        catch { return path.Replace('\\', '/'); }
    }

    private static string? ComputeFileHash(string path) {
        try {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) return ComputeHash(stream);
        } catch { return null; }
    }

    private static string? ComputeStreamHash(Stream stream) {
        if (!stream.CanSeek) return null;
        long position = stream.Position;
        try {
            stream.Position = 0;
            return ComputeHash(stream);
        } finally {
            stream.Position = position;
        }
    }

    private static string? ComputePayloadHash(OneNoteBinaryPayload payload, long maxBytes) {
        if (payload.Length.HasValue && payload.Length.Value > maxBytes) return null;
        using (Stream stream = payload.OpenRead()) return ComputeHashBounded(stream, maxBytes);
    }

    private static string? ComputeHashBounded(Stream stream, long maxBytes) {
        using (var sha = SHA256.Create()) {
            var buffer = new byte[64 * 1024];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total += read;
                if (total > maxBytes) return null;
                sha.TransformBlock(buffer, 0, read, null, 0);
            }
            sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return ToHex(sha.Hash!);
        }
    }

    private static string ComputeHash(string value) => ComputeHash(Encoding.UTF8.GetBytes(value));

    private static string ComputeHash(byte[] bytes) {
        using (var sha = SHA256.Create()) return ToHex(sha.ComputeHash(bytes));
    }

    private static string ComputeHash(Stream stream) {
        using (var sha = SHA256.Create()) return ToHex(sha.ComputeHash(stream));
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        foreach (byte value in bytes) builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private static int EstimateTokenCount(string? value) {
        return string.IsNullOrEmpty(value) ? 0 : Math.Max(1, (value!.Length + 3) / 4);
    }
}
