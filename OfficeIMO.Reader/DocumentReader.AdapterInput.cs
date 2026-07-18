using System;
using System.Globalization;
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
        long? maxInputBytes = ResolveInitialMaxInputBytes(path, options);
        ReaderInputLimits.EnforceFileSize(path, maxInputBytes);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        ReaderAdapterInputSnapshot snapshot = ReadAdapterInput(
            stream,
            path,
            options,
            cancellationToken,
            maxInputBytes);
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
        return ReadAdapterInput(
            stream,
            logicalName,
            options,
            cancellationToken,
            ResolveStreamMaxInputBytes(logicalName, options,
                stream.CanSeek));
    }

    private static ReaderAdapterInputSnapshot ReadAdapterInput(
        Stream stream,
        string logicalName,
        ReaderOptions options,
        CancellationToken cancellationToken,
        long? maxInputBytes) {
        long? originalPosition = TryGetPosition(stream);
        Stream snapshot = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            maxInputBytes,
            cancellationToken,
            out bool ownsSnapshot);
        try {
            byte[] payload = ReadSnapshotPayload(snapshot, cancellationToken, maxInputBytes);
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

    private static byte[] ReadSnapshotPayload(
        Stream snapshot,
        CancellationToken cancellationToken,
        long? maxInputBytes) {
        cancellationToken.ThrowIfCancellationRequested();
        long remaining = snapshot.Length - snapshot.Position;
        if (remaining < 0) {
            throw new IOException("Adapter input stream position exceeds its length.");
        }
        if (maxInputBytes.HasValue && remaining > maxInputBytes.Value) {
            throw new IOException(
                "Input exceeds MaxInputBytes (" +
                remaining.ToString(CultureInfo.InvariantCulture) + " > " +
                maxInputBytes.Value.ToString(CultureInfo.InvariantCulture) + ").");
        }
        if (remaining > int.MaxValue) {
            throw new IOException("Adapter input exceeds the maximum supported byte-array length.");
        }

        var payload = new byte[(int)remaining];
        int offset = 0;
        while (offset < payload.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = snapshot.Read(payload, offset, payload.Length - offset);
            if (read <= 0) {
                throw new EndOfStreamException("Adapter input ended before its declared length was read.");
            }
            offset += read;
        }
        return payload;
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
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
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
