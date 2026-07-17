using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static string BuildPortableSourceId(string sourceKey) {
        if (sourceKey == null) throw new ArgumentNullException(nameof(sourceKey));
        // External identities are not filesystem paths, so their case must not vary by host OS.
        return "src:" + ComputeSha256Hex(sourceKey);
    }

    internal static void ApplyExternalSourceMetadata(
        OfficeDocumentReadResult result,
        string sourceId,
        DateTime? lastWriteUtc,
        long? lengthBytes,
        bool computeHashes) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (string.IsNullOrWhiteSpace(sourceId)) throw new ArgumentException("Source ID cannot be empty.", nameof(sourceId));

        result.Source ??= new OfficeDocumentSource();
        result.Source.SourceId = sourceId;
        if (lastWriteUtc.HasValue) result.Source.LastWriteUtc = lastWriteUtc.Value.ToUniversalTime();
        if (lengthBytes.HasValue) result.Source.LengthBytes = lengthBytes;

        // Source identity participates in chunk hashes; update both together as one core operation.
        IReadOnlyList<ReaderChunk> chunks = result.Chunks ?? Array.Empty<ReaderChunk>();
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            chunk.SourceId = sourceId;
            if (lastWriteUtc.HasValue) chunk.SourceLastWriteUtc = lastWriteUtc.Value.ToUniversalTime();
            if (lengthBytes.HasValue) chunk.SourceLengthBytes = lengthBytes;
            chunk.ChunkHash = computeHashes ? ComputeChunkHash(chunk) : null;
        }
    }
}
