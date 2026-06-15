namespace OfficeIMO.Reader.Rtf;

public static partial class DocumentReaderRtfExtensions {
    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceMetadata source, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (source == null) throw new ArgumentNullException(nameof(source));

        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeSha256Hex(BuildChunkHashInput(chunk));
        }

        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        string safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static SourceMetadata BuildSourceMetadataFromPath(string path, bool computeHash) {
        var normalizedPath = NormalizePathForId(path);
        var source = new SourceMetadata {
            Path = path,
            SourceId = BuildSourceId(normalizedPath),
            SourceHash = computeHash ? TryComputeFileSha256(path) : null
        };

        try {
            var fileInfo = new FileInfo(path);
            if (fileInfo.Exists) {
                source.LastWriteUtc = fileInfo.LastWriteTimeUtc;
                source.LengthBytes = fileInfo.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        return source;
    }

    private static void UpdateSourceMetadataFromSeekableStream(SourceMetadata source, Stream stream, bool computeHash, long startPosition) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        try {
            if (stream.CanSeek) {
                source.LengthBytes = Math.Max(0L, stream.Length - startPosition);
            }
        } catch {
            // Best-effort metadata.
        }

        if (computeHash) {
            source.SourceHash ??= TryComputeStreamSha256(stream, startPosition);
        }
    }

    private static string NormalizeLogicalSourceName(string? sourceName, string fallback) {
        string? trimmed = sourceName?.Trim();
        return string.IsNullOrWhiteSpace(trimmed) ? fallback : trimmed!;
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        try {
            return Path.GetFullPath(path).Replace('\\', '/');
        } catch {
            return path.Replace('\\', '/');
        }
    }

    private static string BuildSourceId(string sourceKey) {
        string normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar == '\\') {
            normalized = normalized.ToLowerInvariant();
        }

        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string? TryComputeFileSha256(string path) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs);
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream, long startPosition) {
        if (stream == null || !stream.CanSeek) return null;

        long position;
        try {
            position = stream.Position;
        } catch {
            return null;
        }

        try {
            stream.Position = startPosition;
            string hash = ComputeSha256Hex(stream);
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
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty)));
    }

    private static string ComputeSha256Hex(byte[] bytes) {
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(bytes ?? Array.Empty<byte>()));
    }

    private static string ComputeSha256Hex(Stream stream) {
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(stream));
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string BuildChunkHashInput(ReaderChunk chunk) {
        return string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty);
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
