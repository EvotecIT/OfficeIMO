using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Progress event kind emitted during folder ingestion.
/// </summary>
public enum ReaderProgressEventKind {
    /// <summary>
    /// Processing started for a file.
    /// </summary>
    FileStarted = 0,
    /// <summary>
    /// Processing finished successfully for a file.
    /// </summary>
    FileCompleted,
    /// <summary>
    /// Processing skipped for a file (limits, parse errors, or metadata failures).
    /// </summary>
    FileSkipped,
    /// <summary>
    /// Folder ingestion completed.
    /// </summary>
    Completed
}

/// <summary>
/// Progress event payload for folder ingestion.
/// </summary>
public sealed class ReaderProgress {
    /// <summary>
    /// Event kind.
    /// </summary>
    public ReaderProgressEventKind Kind { get; set; }

    /// <summary>
    /// Current file path for file-level events.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Optional source identifier for file-level events.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Optional source hash for file-level events.
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Files considered for ingestion so far (allowed extension scope).
    /// </summary>
    public int FilesScanned { get; set; }

    /// <summary>
    /// Files successfully parsed so far.
    /// </summary>
    public int FilesParsed { get; set; }

    /// <summary>
    /// Files skipped so far.
    /// </summary>
    public int FilesSkipped { get; set; }

    /// <summary>
    /// Total bytes accepted for parsed files so far.
    /// </summary>
    public long BytesRead { get; set; }

    /// <summary>
    /// Total chunks emitted so far.
    /// </summary>
    public int ChunksProduced { get; set; }

    /// <summary>
    /// Optional message for skip/reason summaries.
    /// </summary>
    public string? Message { get; set; }

    /// <summary>
    /// Optional current file size in bytes for file-level events.
    /// </summary>
    public long? CurrentFileBytes { get; set; }

    /// <summary>
    /// Optional current file chunk count for file completion events.
    /// </summary>
    public int? CurrentFileChunks { get; set; }

    /// <summary>
    /// Optional current file last-write timestamp (UTC) for file-level events.
    /// </summary>
    public DateTime? CurrentFileLastWriteUtc { get; set; }
}
