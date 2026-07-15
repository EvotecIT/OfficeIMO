using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static ReaderAdapterInputSnapshot ReadAdapterInput(
        string path,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("File '" + path + "' does not exist.", path);
        ReaderInputLimits.EnforceFileSize(path, options.MaxInputBytes);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        ReaderAdapterInputSnapshot snapshot = ReadAdapterInput(stream, path, options, cancellationToken);
        DateTime? lastWriteUtc = null;
        try {
            lastWriteUtc = File.GetLastWriteTimeUtc(path);
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }

        snapshot.Source.Path = path;
        snapshot.Source.SourceId = BuildSourceId(NormalizePathForId(path));
        snapshot.Source.LastWriteUtc = lastWriteUtc;
        return snapshot;
    }

    internal static ReaderAdapterInputSnapshot ReadAdapterInput(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : sourceName!.Trim();
        long? originalPosition = TryGetPosition(stream);
        Stream snapshot = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            options.MaxInputBytes,
            cancellationToken,
            out bool ownsSnapshot);
        try {
            using MemoryStream bytes = CopyToMemory(snapshot, cancellationToken, options.MaxInputBytes);
            byte[] payload = bytes.ToArray();
            return new ReaderAdapterInputSnapshot(
                payload,
                new OfficeDocumentSource {
                    Path = logicalName,
                    SourceId = BuildSourceId(logicalName),
                    SourceHash = options.ComputeHashes ? OfficeDocumentAssetHash.ComputeSha256Hex(payload) : null,
                    LengthBytes = payload.LongLength
                });
        } finally {
            if (ownsSnapshot) {
                snapshot.Dispose();
            } else if (originalPosition.HasValue) {
                TrySetPosition(stream, originalPosition.Value);
            }
        }
    }

    internal static ReaderChunk ApplyAdapterSource(
        ReaderChunk chunk,
        ReaderAdapterInputSnapshot input,
        bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (input == null) throw new ArgumentNullException(nameof(input));
        chunk.Location ??= new ReaderLocation();
        chunk.Location.Path ??= input.Source.Path;
        chunk.SourceId = input.Source.SourceId;
        chunk.SourceHash = computeHashes ? input.Source.SourceHash : null;
        chunk.SourceLastWriteUtc = input.Source.LastWriteUtc;
        chunk.SourceLengthBytes = input.Source.LengthBytes;
        chunk.ChunkHash = computeHashes ? ComputeChunkHash(chunk) : null;
        return chunk;
    }

    private static long? TryGetPosition(Stream stream) {
        if (!stream.CanSeek) return null;
        try {
            return stream.Position;
        } catch (NotSupportedException) {
            return null;
        }
    }

    private static void TrySetPosition(Stream stream, long position) {
        try {
            stream.Position = position;
        } catch (IOException) {
        } catch (NotSupportedException) {
        }
    }
}

internal sealed class ReaderAdapterInputSnapshot {
    internal ReaderAdapterInputSnapshot(byte[] bytes, OfficeDocumentSource source) {
        Bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
        Source = source ?? throw new ArgumentNullException(nameof(source));
    }

    internal byte[] Bytes { get; }

    internal OfficeDocumentSource Source { get; }
}
