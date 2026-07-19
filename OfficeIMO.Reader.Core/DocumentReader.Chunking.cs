using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceInfo source, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (source == null) throw new ArgumentNullException(nameof(source));
        chunk.Location ??= new ReaderLocation();
        chunk.Location.Path ??= source.Path;
        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) chunk.ChunkHash = ComputeChunkHash(chunk);
        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        int length = text?.Length ?? 0;
        return length == 0 ? 0 : Math.Max(1, (length + 3) / 4);
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        string data = string.Join("|",
            chunk.Kind.ToString(), chunk.SourceId ?? string.Empty,
            chunk.Location?.Path ?? string.Empty, chunk.Location?.HeadingPath ?? string.Empty,
            chunk.Location?.HeadingSlug ?? string.Empty, chunk.Location?.SourceBlockKind ?? string.Empty,
            chunk.Location?.BlockAnchor ?? string.Empty, chunk.Location?.Sheet ?? string.Empty,
            chunk.Location?.A1Range ?? string.Empty,
            chunk.Location?.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location?.Slide?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location?.StartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty, chunk.Markdown ?? string.Empty);
        return ComputeSha256Hex(data);
    }

    private static SourceInfo BuildSourceInfoFromPath(string path, bool computeHash, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var source = new SourceInfo {
            Path = path,
            SourceId = BuildSourceId(NormalizePathForId(path))
        };
        try {
            var info = new FileInfo(path);
            if (info.Exists) {
                source.LastWriteUtc = info.LastWriteTimeUtc;
                source.LengthBytes = info.Length;
            }
        } catch {
            // Source metadata is best effort.
        }
        if (computeHash) source.SourceHash = TryComputeFileSha256(path, cancellationToken);
        return source;
    }

    private static SourceInfo BuildSourceInfoFromStream(Stream stream, string? sourceName, bool computeHash, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : sourceName!.Trim();
        var source = new SourceInfo { Path = logicalName, SourceId = BuildSourceId(logicalName) };
        try {
            if (stream.CanSeek) source.LengthBytes = stream.Length;
        } catch {
            // Source metadata is best effort.
        }
        if (computeHash) source.SourceHash = TryComputeStreamSha256(stream, cancellationToken);
        return source;
    }

    private static ReaderSourceDocument BuildSourceDocument(SourceInfo source, bool parsed, IReadOnlyList<ReaderChunk>? chunks, IReadOnlyList<string>? sourceWarnings) {
        IReadOnlyList<ReaderChunk> materialized = chunks ?? Array.Empty<ReaderChunk>();
        var warnings = new List<string>();
        if (sourceWarnings != null) foreach (string warning in sourceWarnings) AddWarning(warnings, warning);
        int tokenEstimate = 0;
        foreach (ReaderChunk chunk in materialized) {
            tokenEstimate += chunk.TokenEstimate ?? EstimateTokenCount(chunk.Text);
            if (chunk.Warnings != null) foreach (string warning in chunk.Warnings) AddWarning(warnings, warning);
        }
        return new ReaderSourceDocument {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            SourceLastWriteUtc = source.LastWriteUtc,
            SourceLengthBytes = source.LengthBytes,
            Parsed = parsed,
            ChunksProduced = materialized.Count,
            TokenEstimateTotal = tokenEstimate,
            Warnings = warnings.Count == 0 ? null : warnings,
            Chunks = materialized
        };
    }

    private static ReaderChunk BuildFolderWarningChunk(string path, int warningIndex, string warning) {
        string fileName = Path.GetFileName(path);
        if (string.IsNullOrWhiteSpace(fileName)) fileName = "folder";
        return new ReaderChunk {
            Id = "warn:" + fileName + ":" + warningIndex.ToString("D4", CultureInfo.InvariantCulture),
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation { Path = path, BlockIndex = warningIndex },
            Text = string.Empty,
            Warnings = new[] { warning }
        };
    }

    private static void AddWarning(List<string> warnings, string? warning) {
        if (string.IsNullOrWhiteSpace(warning)) return;
        if (!warnings.Any(existing => string.Equals(existing, warning, StringComparison.OrdinalIgnoreCase))) warnings.Add(warning!);
    }

    private static string? TryComputeFileSha256(string path, CancellationToken cancellationToken) {
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(stream, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return null;
        long position;
        try { position = stream.Position; } catch { return null; }
        try {
            stream.Position = 0;
            return ComputeSha256Hex(stream, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return null;
        } finally {
            try { stream.Position = position; } catch { }
        }
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = SHA256.Create();
        return ConvertToHexLower(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty)));
    }

    private static string ComputeSha256Hex(Stream stream, CancellationToken cancellationToken = default) {
        using var sha = SHA256.Create();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            sha.TransformBlock(buffer, 0, read, buffer, 0);
        }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return ConvertToHexLower(sha.Hash!);
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var result = new StringBuilder(bytes.Length * 2);
        foreach (byte value in bytes) result.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        return result.ToString();
    }

    private static string BuildSourceId(string sourceKey) {
        string normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar == '\\') normalized = normalized.ToLowerInvariant();
        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        string value;
        try { value = Path.GetFullPath(path); } catch { value = path; }
        return value.Replace('\\', '/');
    }
}
