using System.Runtime.InteropServices;
using System.Globalization;

namespace OfficeIMO.Reader.Zip {

    internal static partial class ZipReaderAdapter {
        private static ReaderChunk BuildWarningChunk(
            ArchiveSourceMetadata archiveSource,
            string archivePath,
            string entryName,
            int warningIndex,
            string warning,
            bool computeHashes,
            long? sourceLengthBytes = null,
            DateTimeOffset? sourceLastWriteTime = null) {
            var warningPath = string.IsNullOrWhiteSpace(entryName)
                ? archivePath
                : BuildVirtualPath(archivePath, entryName);

            var chunk = new ReaderChunk {
                Id = $"zip-warning-{warningIndex.ToString("D4", System.Globalization.CultureInfo.InvariantCulture)}",
                Kind = ReaderInputKind.Zip,
                Location = new ReaderLocation {
                    Path = warningPath,
                    BlockIndex = warningIndex
                },
                Text = warning,
                Warnings = new[] { warning }
            };

            ApplyWarningSourceMetadata(chunk, archiveSource, warningPath, sourceLengthBytes, sourceLastWriteTime, computeHashes);
            return chunk;
        }

        private static string BuildVirtualPath(string zipPath, string entryName) {
            return zipPath + "::" + entryName.Replace('\\', '/');
        }

        private static void ApplyVirtualSourceMetadata(
            ReaderChunk chunk,
            ArchiveSourceMetadata archiveSource,
            string virtualPath,
            long uncompressedLength,
            DateTimeOffset lastWriteTime,
            bool computeHashes) {
            if (chunk == null) throw new ArgumentNullException(nameof(chunk));

            chunk.Location.Path = virtualPath;
            chunk.SourceId = BuildSourceId(virtualPath);
            chunk.SourceHash ??= archiveSource.SourceHash;
            chunk.SourceLengthBytes = uncompressedLength >= 0 ? uncompressedLength : chunk.SourceLengthBytes;

            var sourceLastWriteUtc = NormalizeLastWriteUtc(lastWriteTime);
            if (sourceLastWriteUtc.HasValue) {
                chunk.SourceLastWriteUtc = sourceLastWriteUtc;
            }
            chunk.SourceLastWriteUtc ??= archiveSource.LastWriteUtc;

            if (!chunk.TokenEstimate.HasValue) {
                chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
            }

            if (computeHashes) {
                chunk.ChunkHash = ComputeChunkHash(chunk);
            }
        }

        private static void ApplyWarningSourceMetadata(
            ReaderChunk chunk,
            ArchiveSourceMetadata archiveSource,
            string sourcePath,
            long? sourceLengthBytes,
            DateTimeOffset? sourceLastWriteTime,
            bool computeHashes) {
            if (chunk == null) throw new ArgumentNullException(nameof(chunk));

            chunk.Location.Path = sourcePath;
            chunk.SourceId = BuildSourceId(sourcePath);
            chunk.SourceHash ??= archiveSource.SourceHash;

            if (sourceLengthBytes.HasValue && sourceLengthBytes.Value >= 0) {
                chunk.SourceLengthBytes = sourceLengthBytes.Value;
            } else if (string.Equals(sourcePath, archiveSource.Path, StringComparison.OrdinalIgnoreCase)) {
                chunk.SourceLengthBytes ??= archiveSource.LengthBytes;
            }

            if (sourceLastWriteTime.HasValue) {
                var normalized = NormalizeLastWriteUtc(sourceLastWriteTime.Value);
                if (normalized.HasValue) {
                    chunk.SourceLastWriteUtc = normalized.Value;
                }
            } else if (string.Equals(sourcePath, archiveSource.Path, StringComparison.OrdinalIgnoreCase)) {
                chunk.SourceLastWriteUtc ??= archiveSource.LastWriteUtc;
            }

            if (!chunk.TokenEstimate.HasValue) {
                chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
            }

            if (computeHashes) {
                chunk.ChunkHash = ComputeChunkHash(chunk);
            }
        }

        private static int EstimateTokenCount(string? text) {
            var safeText = text ?? string.Empty;
            if (safeText.Length == 0) return 0;
            return Math.Max(1, (safeText.Length + 3) / 4);
        }

        private static ArchiveSourceMetadata BuildArchiveSourceMetadataFromPath(string zipPath, Stream archiveStream, bool computeHash) {
            var normalizedPath = NormalizePathForVirtualPath(zipPath);
            DateTime? lastWriteUtc = null;
            long? lengthBytes = null;
            try {
                var fileInfo = new FileInfo(zipPath);
                if (fileInfo.Exists) {
                    lastWriteUtc = fileInfo.LastWriteTimeUtc;
                }
            } catch {
                // Best-effort metadata.
            }
            try {
                lengthBytes = archiveStream.Length;
            } catch {
                // Best-effort metadata.
            }

            return new ArchiveSourceMetadata {
                Path = normalizedPath,
                SourceHash = computeHash ? TryComputeStreamSha256(archiveStream) : null,
                LastWriteUtc = lastWriteUtc,
                LengthBytes = lengthBytes
            };
        }

        private static ArchiveSourceMetadata BuildArchiveSourceMetadataFromStream(Stream stream, string sourceName, bool computeHash) {
            long? lengthBytes = null;
            try {
                if (stream.CanSeek) {
                    lengthBytes = stream.Length;
                }
            } catch {
                // Best-effort metadata.
            }

            return new ArchiveSourceMetadata {
                Path = sourceName,
                SourceHash = computeHash ? TryComputeStreamSha256(stream) : null,
                LastWriteUtc = null,
                LengthBytes = lengthBytes
            };
        }

        private static DateTime? NormalizeLastWriteUtc(DateTimeOffset lastWriteTime) {
            if (lastWriteTime == default) return null;

            // Zip archives sometimes surface sentinel timestamps; keep only practical values.
            var utc = lastWriteTime.UtcDateTime;
            if (utc.Year < 1980) return null;
            return utc;
        }

        private static string BuildSourceId(string sourceKey) {
            var normalized = NormalizeSourceKeyForId(sourceKey);

            return "src:" + ComputeSha256Hex(normalized);
        }

        private static string NormalizePathForVirtualPath(string path) {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            string fullPath;
            try {
                fullPath = Path.GetFullPath(path);
            } catch {
                fullPath = path;
            }

            fullPath = ResolveExistingFullPath(fullPath);
            return fullPath.Replace('\\', '/');
        }

        private static string ResolveExistingFullPath(string fullPath) {
            if (Path.DirectorySeparatorChar == '\\') {
                return fullPath;
            }

            IntPtr resolvedPathPointer = IntPtr.Zero;
            try {
                resolvedPathPointer = UnixRealPath(fullPath, IntPtr.Zero);
                if (resolvedPathPointer == IntPtr.Zero) {
                    return fullPath;
                }

                var resolvedPath = Marshal.PtrToStringAnsi(resolvedPathPointer);
                return string.IsNullOrWhiteSpace(resolvedPath) ? fullPath : resolvedPath;
            } catch {
                return fullPath;
            } finally {
                if (resolvedPathPointer != IntPtr.Zero) {
                    UnixFree(resolvedPathPointer);
                }
            }
        }

        [DllImport("libc", EntryPoint = "realpath", CharSet = CharSet.Ansi)]
        private static extern IntPtr UnixRealPath(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "free")]
        private static extern void UnixFree(IntPtr pointer);

        private static string NormalizeSourceKeyForId(string? sourceKey) {
            var normalized = sourceKey ?? string.Empty;
            if (Path.DirectorySeparatorChar != '\\') {
                return normalized;
            }

            var virtualPathSeparatorIndex = normalized.IndexOf("::", StringComparison.Ordinal);
            if (virtualPathSeparatorIndex < 0) {
                return normalized.ToLowerInvariant();
            }

            var archivePath = normalized.Substring(0, virtualPathSeparatorIndex).ToLowerInvariant();
            var archiveEntryPath = normalized.Substring(virtualPathSeparatorIndex);
            return archivePath + archiveEntryPath;
        }

        private static string ComputeChunkHash(ReaderChunk chunk) => DocumentReaderEngine.ComputeChunkHash(chunk);

        private static string ComputeSha256Hex(string value) {
            using var sha = System.Security.Cryptography.SHA256.Create();
            var bytes = System.Text.Encoding.UTF8.GetBytes(value ?? string.Empty);
            var hash = sha.ComputeHash(bytes);

            var sb = new System.Text.StringBuilder(hash.Length * 2);
            for (int i = 0; i < hash.Length; i++) {
                sb.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static string ComputeSha256Hex(Stream stream) {
            using var sha = System.Security.Cryptography.SHA256.Create();
            var hash = sha.ComputeHash(stream);

            var sb = new System.Text.StringBuilder(hash.Length * 2);
            for (int i = 0; i < hash.Length; i++) {
                sb.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static string? TryComputeStreamSha256(Stream stream) {
            if (stream == null || !stream.CanSeek) return null;

            long position;
            try {
                position = stream.Position;
            } catch {
                return null;
            }

            try {
                stream.Position = 0;
                var hash = ComputeSha256Hex(stream);
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

        private sealed class WarningCounter {
            public int Value { get; set; }
            public int Next() {
                var current = Value;
                Value = current + 1;
                return current;
            }
        }

        private sealed class ArchiveSourceMetadata {
            public string Path { get; set; } = string.Empty;
            public string? SourceHash { get; set; }
            public DateTime? LastWriteUtc { get; set; }
            public long? LengthBytes { get; set; }
        }
    }
}
