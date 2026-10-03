using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

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
    // Path handlers may reopen the source. Reject a changing source rather than combine content
    // from one revision with another revision's hash. Incremental callers validate between pulls.
    private static void ValidateUnchangedPathSource(SourceInfo source, CancellationToken token, bool verifyHash = true) {
        SourceInfo current = BuildSourceInfoFromPath(source.Path, verifyHash && source.SourceHash != null, token);
        ValidatePathSourceRevision(source, current, verifyHash);
    }

    private static void ValidatePathSourceRevision(SourceInfo source, SourceInfo current, bool verifyHash) {
        if (source.LengthBytes != current.LengthBytes || source.LastWriteUtc != current.LastWriteUtc ||
            (verifyHash && source.SourceHash != null && source.SourceHash != current.SourceHash))
            throw new IOException("Source changed while Reader was extracting '" + source.Path + "'. Retry with a stable source.");
    }

}
