using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Per-file ingestion status summary.
/// </summary>
public sealed class ReaderIngestFileResult {
    /// <summary>
    /// Source file path.
    /// </summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>
    /// Stable source identifier.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Optional source content hash.
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Optional source last-write timestamp.
    /// </summary>
    public DateTime? SourceLastWriteUtc { get; set; }

    /// <summary>
    /// Optional source length in bytes.
    /// </summary>
    public long? SourceLengthBytes { get; set; }

    /// <summary>
    /// True when parsing succeeded for this file.
    /// </summary>
    public bool Parsed { get; set; }

    /// <summary>
    /// Number of chunks emitted for this file.
    /// </summary>
    public int ChunksProduced { get; set; }

    /// <summary>
    /// Optional warnings associated with this file.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Per-source ingestion payload optimized for direct database upserts.
/// </summary>
public sealed class ReaderSourceDocument {
    /// <summary>
    /// Source file path.
    /// </summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>
    /// Stable source identifier.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Optional source content hash.
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Optional source last-write timestamp.
    /// </summary>
    public DateTime? SourceLastWriteUtc { get; set; }

    /// <summary>
    /// Optional source length in bytes.
    /// </summary>
    public long? SourceLengthBytes { get; set; }

    /// <summary>
    /// True when parsing succeeded for this source.
    /// </summary>
    public bool Parsed { get; set; }

    /// <summary>
    /// Number of chunks emitted for this source.
    /// </summary>
    public int ChunksProduced { get; set; }

    /// <summary>
    /// Aggregated token estimate across emitted chunks.
    /// </summary>
    public int TokenEstimateTotal { get; set; }

    /// <summary>
    /// Optional source-level warnings (parse errors, limit skips, extraction warnings).
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }

    /// <summary>
    /// Emitted chunks for this source (empty for skipped files).
    /// </summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; set; } = Array.Empty<ReaderChunk>();
}

/// <summary>
/// Detailed document-oriented read result for a single file or a folder path.
/// </summary>
public sealed class ReaderPathDocumentResult {
    /// <summary>
    /// Source file paths included in the result.
    /// </summary>
    public IReadOnlyList<string> Files { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Source-level documents returned for the path.
    /// </summary>
    public IReadOnlyList<ReaderSourceDocument> Documents { get; set; } = Array.Empty<ReaderSourceDocument>();

    /// <summary>
    /// Files considered for ingestion (allowed extension scope).
    /// </summary>
    public int FilesScanned { get; set; }

    /// <summary>
    /// Files parsed successfully.
    /// </summary>
    public int FilesParsed { get; set; }

    /// <summary>
    /// Files skipped.
    /// </summary>
    public int FilesSkipped { get; set; }

    /// <summary>
    /// Bytes accepted for parsed files.
    /// </summary>
    public long BytesRead { get; set; }

    /// <summary>
    /// Total chunks produced before any returned-chunk shaping.
    /// </summary>
    public int ChunksProduced { get; set; }

    /// <summary>
    /// Total chunk objects materialized in <see cref="Documents"/>.
    /// </summary>
    public int ChunksReturned { get; set; }

    /// <summary>
    /// Aggregated token estimate across returned chunks.
    /// </summary>
    public int TokenEstimateReturned { get; set; }

    /// <summary>
    /// True when returned chunk materialization was truncated by caller limits.
    /// </summary>
    public bool Truncated { get; set; }

    /// <summary>
    /// Aggregated warnings associated with the read.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Detailed folder-ingestion result optimized for indexing pipelines.
/// </summary>
public sealed class ReaderIngestResult {
    /// <summary>
    /// File-level statuses emitted during ingestion.
    /// </summary>
    public IReadOnlyList<ReaderIngestFileResult> Files { get; set; } = Array.Empty<ReaderIngestFileResult>();

    /// <summary>
    /// Emitted chunks when requested by the caller.
    /// </summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; set; } = Array.Empty<ReaderChunk>();

    /// <summary>
    /// Files considered for ingestion (allowed extension scope).
    /// </summary>
    public int FilesScanned { get; set; }

    /// <summary>
    /// Files parsed successfully.
    /// </summary>
    public int FilesParsed { get; set; }

    /// <summary>
    /// Files skipped.
    /// </summary>
    public int FilesSkipped { get; set; }

    /// <summary>
    /// Bytes accepted for parsed files.
    /// </summary>
    public long BytesRead { get; set; }

    /// <summary>
    /// Total chunks produced.
    /// </summary>
    public int ChunksProduced { get; set; }

    /// <summary>
    /// Aggregated ingestion warnings.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}
