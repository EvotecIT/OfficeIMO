using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Creates the shared v5 document envelope from adapter-produced chunks and optional source data.
    /// </summary>
    /// <remarks>
    /// Format adapters can use this method for the common chunk, table, visual, page, metadata,
    /// diagnostic, and OCR-candidate projection before adding format-owned rich structures.
    /// </remarks>
    public static OfficeDocumentReadResult CreateDocumentResult(
        IEnumerable<ReaderChunk> chunks,
        ReaderInputKind fallbackKind,
        OfficeDocumentSource? source = null,
        IEnumerable<string>? capabilities = null,
        IReadOnlyList<OfficeDocumentAsset>? assets = null) =>
        CreateDocumentResult(chunks, fallbackKind, source, capabilities, assets, ocrCandidates: null);

    /// <summary>
    /// Creates the shared v5 document envelope with adapter-produced OCR candidates.
    /// </summary>
    public static OfficeDocumentReadResult CreateDocumentResult(
        IEnumerable<ReaderChunk> chunks,
        ReaderInputKind fallbackKind,
        OfficeDocumentSource? source,
        IEnumerable<string>? capabilities,
        IReadOnlyList<OfficeDocumentAsset>? assets,
        IReadOnlyList<OfficeDocumentOcrCandidate>? ocrCandidates) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));
        ReaderChunk[] materializedChunks = chunks.ToArray();
        ReaderChunk? first = materializedChunks.Length == 0 ? null : materializedChunks[0];
        OfficeDocumentSource effectiveSource = source ?? new OfficeDocumentSource {
            Path = first?.Location.Path,
            SourceId = first?.SourceId,
            SourceHash = first?.SourceHash,
            LastWriteUtc = first?.SourceLastWriteUtc,
            LengthBytes = first?.SourceLengthBytes
        };
        string sourceName = effectiveSource.Path ?? first?.Location.Path ?? "memory";
        OfficeDocumentReadResult result = BuildChunkDocumentResult(
            materializedChunks,
            sourceName,
            fallbackKind,
            effectiveSource,
            assets,
            ocrCandidates: ocrCandidates);
        if (capabilities != null) {
            result.CapabilitiesUsed = result.CapabilitiesUsed
                .Concat(capabilities)
                .Where(static capability => !string.IsNullOrWhiteSpace(capability))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }
        return result;
    }
}
