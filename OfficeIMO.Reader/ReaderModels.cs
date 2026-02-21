using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// The detected input kind used by <see cref="DocumentReader"/>.
/// </summary>
public enum ReaderInputKind {
    /// <summary>
    /// Unknown/other file kind treated as plain text.
    /// </summary>
    Unknown = 0,
    /// <summary>
    /// Word document (DOCX/DOCM).
    /// </summary>
    Word,
    /// <summary>
    /// Excel workbook (XLSX/XLSM).
    /// </summary>
    Excel,
    /// <summary>
    /// PowerPoint presentation (PPTX/PPTM).
    /// </summary>
    PowerPoint,
    /// <summary>
    /// Markdown text file.
    /// </summary>
    Markdown,
    /// <summary>
    /// Generic text file.
    /// </summary>
    Text,
    /// <summary>
    /// PDF document.
    /// </summary>
    Pdf,
    /// <summary>
    /// CSV/TSV structured text.
    /// </summary>
    Csv,
    /// <summary>
    /// JSON structured text.
    /// </summary>
    Json,
    /// <summary>
    /// XML structured text.
    /// </summary>
    Xml,
    /// <summary>
    /// HTML document.
    /// </summary>
    Html,
    /// <summary>
    /// ZIP archive.
    /// </summary>
    Zip,
    /// <summary>
    /// EPUB e-book archive.
    /// </summary>
    Epub
}

/// <summary>
/// A normalized extraction chunk produced by <see cref="DocumentReader"/>.
/// </summary>
public sealed class ReaderChunk {
    /// <summary>
    /// Stable, ASCII-only identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// The kind of input that produced this chunk.
    /// </summary>
    public ReaderInputKind Kind { get; set; }

    /// <summary>
    /// Source location information for citations and debugging.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Stable identifier for the source document.
    /// For file-based reads this is deterministic for a given normalized path.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Optional content hash for the source document (for incremental upserts).
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Optional content hash for this chunk (for incremental upserts).
    /// </summary>
    public string? ChunkHash { get; set; }

    /// <summary>
    /// Optional source last-write timestamp (UTC) when available.
    /// </summary>
    public DateTime? SourceLastWriteUtc { get; set; }

    /// <summary>
    /// Optional source length in bytes when available.
    /// </summary>
    public long? SourceLengthBytes { get; set; }

    /// <summary>
    /// Estimated token count (best-effort heuristic) for prompt budgeting.
    /// </summary>
    public int? TokenEstimate { get; set; }

    /// <summary>
    /// Plain text representation of the chunk.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional Markdown representation of the chunk.
    /// </summary>
    public string? Markdown { get; set; }

    /// <summary>
    /// Optional structured tables extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderTable>? Tables { get; set; }

    /// <summary>
    /// Optional warnings about truncation or unsupported content.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Generic location metadata used across supported formats.
/// </summary>
public sealed class ReaderLocation {
    /// <summary>
    /// Source path used for citations.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Optional emitted chunk index (0-based) in the order produced by <see cref="DocumentReader"/>.
    /// </summary>
    public int? BlockIndex { get; set; }

    /// <summary>
    /// Optional source block index within the input document (producer-defined).
    /// For Word, this is the first block index included in the emitted chunk.
    /// </summary>
    public int? SourceBlockIndex { get; set; }

    /// <summary>
    /// Optional 1-based start line number (Markdown/text inputs).
    /// </summary>
    public int? StartLine { get; set; }

    /// <summary>
    /// Optional heading path label (for example "H1 &gt; H2").
    /// </summary>
    public string? HeadingPath { get; set; }

    /// <summary>
    /// Optional sheet name (Excel).
    /// </summary>
    public string? Sheet { get; set; }

    /// <summary>
    /// Optional A1 range descriptor (Excel).
    /// </summary>
    public string? A1Range { get; set; }

    /// <summary>
    /// Optional 1-based slide number (PowerPoint).
    /// </summary>
    public int? Slide { get; set; }

    /// <summary>
    /// Optional 1-based page number (PDF).
    /// </summary>
    public int? Page { get; set; }
}

/// <summary>
/// Minimal table model for ingestion (columns + rows).
/// </summary>
public sealed class ReaderTable {
    /// <summary>
    /// Optional title/label.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Column headers.
    /// </summary>
    public IReadOnlyList<string> Columns { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Rows aligned with <see cref="Columns"/>.
    /// </summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; set; } = Array.Empty<IReadOnlyList<string>>();

    /// <summary>
    /// Total row count before any truncation.
    /// </summary>
    public int TotalRowCount { get; set; }

    /// <summary>
    /// True when <see cref="Rows"/> was truncated compared to <see cref="TotalRowCount"/>.
    /// </summary>
    public bool Truncated { get; set; }
}

/// <summary>
/// Custom handler registration model for extending <see cref="DocumentReader"/> without hard dependencies.
/// </summary>
public sealed class ReaderHandlerRegistration {
    /// <summary>
    /// Stable unique identifier for this handler (for example: "officeimo.reader.epub").
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Optional display name shown in capability listings.
    /// </summary>
    public string? DisplayName { get; set; }

    /// <summary>
    /// Optional handler description shown in capability listings.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Kind advertised by this handler for detect/capability workflows.
    /// </summary>
    public ReaderInputKind Kind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>
    /// File extensions handled by this registration (for example: ".epub", ".zip").
    /// </summary>
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Path-based reader delegate.
    /// </summary>
    public Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadPath { get; set; }

    /// <summary>
    /// Stream-based reader delegate.
    /// </summary>
    public Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadStream { get; set; }

    /// <summary>
    /// Optional advertised default max input bytes for this handler.
    /// Null means "no handler-specific default advertised".
    /// </summary>
    public long? DefaultMaxInputBytes { get; set; }

    /// <summary>
    /// Advertised warning model for this handler.
    /// </summary>
    public ReaderWarningBehavior WarningBehavior { get; set; } = ReaderWarningBehavior.Mixed;

    /// <summary>
    /// True when this handler advertises deterministic chunk ordering/output for identical input.
    /// </summary>
    public bool DeterministicOutput { get; set; } = true;
}

/// <summary>
/// Immutable capability descriptor for built-in and registered handlers.
/// </summary>
public sealed class ReaderHandlerCapability {
    /// <summary>
    /// Stable unique handler identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable name.
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Optional handler description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Advertised input kind.
    /// </summary>
    public ReaderInputKind Kind { get; set; }

    /// <summary>
    /// Extensions served by this handler.
    /// </summary>
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();

    /// <summary>
    /// True for built-in reader handlers.
    /// </summary>
    public bool IsBuiltIn { get; set; }

    /// <summary>
    /// True when path-based read delegate is available.
    /// </summary>
    public bool SupportsPath { get; set; }

    /// <summary>
    /// True when stream-based read delegate is available.
    /// </summary>
    public bool SupportsStream { get; set; }

    /// <summary>
    /// Capability schema identifier for host integration contracts.
    /// </summary>
    public string SchemaId { get; set; } = ReaderCapabilitySchema.Id;

    /// <summary>
    /// Capability schema version for host integration contracts.
    /// </summary>
    public int SchemaVersion { get; set; } = ReaderCapabilitySchema.Version;

    /// <summary>
    /// Optional advertised default max input bytes for this handler.
    /// Null means no handler-specific default is advertised.
    /// </summary>
    public long? DefaultMaxInputBytes { get; set; }

    /// <summary>
    /// Advertised warning model for this handler.
    /// </summary>
    public ReaderWarningBehavior WarningBehavior { get; set; } = ReaderWarningBehavior.Mixed;

    /// <summary>
    /// True when this handler advertises deterministic chunk ordering/output for identical input.
    /// </summary>
    public bool DeterministicOutput { get; set; } = true;
}

/// <summary>
/// Stable capability schema contract values exposed by <see cref="DocumentReader.GetCapabilities(bool, bool)"/>.
/// </summary>
public static class ReaderCapabilitySchema {
    /// <summary>
    /// Stable schema identifier.
    /// </summary>
    public const string Id = "officeimo.reader.capability";

    /// <summary>
    /// Current schema version.
    /// </summary>
    public const int Version = 1;
}

/// <summary>
/// Advertised warning behavior model for reader handlers.
/// </summary>
public enum ReaderWarningBehavior {
    /// <summary>
    /// Handler may both emit warning chunks and throw exceptions, depending on scenario.
    /// </summary>
    Mixed = 0,
    /// <summary>
    /// Handler prefers warning chunks over throwing for recoverable issues.
    /// </summary>
    WarningChunksOnly = 1,
    /// <summary>
    /// Handler prefers exception-based signaling for issues.
    /// </summary>
    ExceptionsOnly = 2
}

/// <summary>
/// Machine-readable capability manifest for host discovery/integration.
/// </summary>
public sealed class ReaderCapabilityManifest {
    /// <summary>
    /// Capability schema identifier.
    /// </summary>
    public string SchemaId { get; set; } = ReaderCapabilitySchema.Id;

    /// <summary>
    /// Capability schema version.
    /// </summary>
    public int SchemaVersion { get; set; } = ReaderCapabilitySchema.Version;

    /// <summary>
    /// Discovered handler capabilities included in this manifest.
    /// </summary>
    public IReadOnlyList<ReaderHandlerCapability> Handlers { get; set; } = Array.Empty<ReaderHandlerCapability>();
}

/// <summary>
/// Options for host bootstrap workflows that auto-register modular handlers and emit capability manifests.
/// </summary>
public sealed class ReaderHostBootstrapOptions {
    /// <summary>
    /// When true, discovered modular registrars can replace conflicting existing custom handlers.
    /// Default: true.
    /// </summary>
    public bool ReplaceExistingHandlers { get; set; } = true;

    /// <summary>
    /// When true, include built-in handlers in the returned capability manifest. Default: true.
    /// </summary>
    public bool IncludeBuiltInCapabilities { get; set; } = true;

    /// <summary>
    /// When true, include custom handlers in the returned capability manifest. Default: true.
    /// </summary>
    public bool IncludeCustomCapabilities { get; set; } = true;

    /// <summary>
    /// When true, indents the returned manifest JSON payload. Default: false.
    /// </summary>
    public bool IndentedManifestJson { get; set; }
}

/// <summary>
/// Output payload for host bootstrap workflows.
/// </summary>
public sealed class ReaderHostBootstrapResult {
    /// <summary>
    /// Prefix used for loaded-assembly bootstrap discovery, when applicable.
    /// </summary>
    public string? AssemblyNamePrefix { get; set; }

    /// <summary>
    /// Effective replace-existing behavior used for registrar invocation.
    /// </summary>
    public bool ReplaceExistingHandlers { get; set; }

    /// <summary>
    /// Registrars that were discovered and invoked during bootstrap.
    /// </summary>
    public IReadOnlyList<ReaderHandlerRegistrarDescriptor> RegisteredHandlers { get; set; } = Array.Empty<ReaderHandlerRegistrarDescriptor>();

    /// <summary>
    /// Capability manifest produced after registration.
    /// </summary>
    public ReaderCapabilityManifest Manifest { get; set; } = new ReaderCapabilityManifest();

    /// <summary>
    /// JSON representation of <see cref="Manifest"/> for host transport.
    /// </summary>
    public string ManifestJson { get; set; } = "{}";
}

/// <summary>
/// Descriptor for a discoverable modular handler registrar method.
/// </summary>
public sealed class ReaderHandlerRegistrarDescriptor {
    /// <summary>
    /// Handler identifier declared by the registrar.
    /// </summary>
    public string HandlerId { get; set; } = string.Empty;

    /// <summary>
    /// Assembly name containing the registrar method.
    /// </summary>
    public string AssemblyName { get; set; } = string.Empty;

    /// <summary>
    /// Fully qualified type name containing the registrar method.
    /// </summary>
    public string TypeName { get; set; } = string.Empty;

    /// <summary>
    /// Registrar method name.
    /// </summary>
    public string MethodName { get; set; } = string.Empty;
}

/// <summary>
/// Marks a static registration method that can be discovered by
/// <see cref="DocumentReader.DiscoverHandlerRegistrars(IEnumerable{System.Reflection.Assembly})"/>.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class ReaderHandlerRegistrarAttribute : Attribute {
    /// <summary>
    /// Creates a registrar attribute for the specified handler identifier.
    /// </summary>
    public ReaderHandlerRegistrarAttribute(string handlerId) {
        HandlerId = handlerId;
    }

    /// <summary>
    /// Handler identifier exposed by the registrar.
    /// </summary>
    public string HandlerId { get; }
}

/// <summary>
/// Options controlling extraction behavior for <see cref="DocumentReader"/>.
/// </summary>
public sealed class ReaderOptions {
    /// <summary>
    /// Optional maximum input size in bytes enforced by <see cref="DocumentReader"/> when reading from a file or seekable stream.
    /// When null, no size limit is enforced.
    /// </summary>
    public long? MaxInputBytes { get; set; }

    /// <summary>
    /// OpenXML security: maximum characters allowed per part when opening OpenXML packages (best-effort).
    /// When null, the OpenXML SDK default is used.
    /// </summary>
    public long? OpenXmlMaxCharactersInPart { get; set; } = 10_000_000;

    /// <summary>
    /// Maximum characters per emitted chunk (best-effort).
    /// </summary>
    public int MaxChars { get; set; } = 8_000;

    /// <summary>
    /// Maximum number of table rows included per table chunk (best-effort).
    /// </summary>
    public int MaxTableRows { get; set; } = 200;

    /// <summary>
    /// When true, include Word footnotes as a final chunk. Default: true.
    /// </summary>
    public bool IncludeWordFootnotes { get; set; } = true;

    /// <summary>
    /// When true, include PowerPoint speaker notes when present. Default: true.
    /// </summary>
    public bool IncludePowerPointNotes { get; set; } = true;

    /// <summary>
    /// Excel: when true, treat the first row as headers. Default: true.
    /// </summary>
    public bool ExcelHeadersInFirstRow { get; set; } = true;

    /// <summary>
    /// Excel: number of worksheet rows per emitted chunk. Default: 200.
    /// </summary>
    public int ExcelChunkRows { get; set; } = 200;

    /// <summary>
    /// Excel: optional sheet name. When null, all sheets are extracted.
    /// </summary>
    public string? ExcelSheetName { get; set; }

    /// <summary>
    /// Excel: optional A1 range. When null, the sheet's used range is used.
    /// </summary>
    public string? ExcelA1Range { get; set; }

    /// <summary>
    /// Markdown: when true, chunk by headings where possible. Default: true.
    /// </summary>
    public bool MarkdownChunkByHeadings { get; set; } = true;

    /// <summary>
    /// When true, computes source/chunk hashes for incremental indexing workflows. Default: true.
    /// </summary>
    public bool ComputeHashes { get; set; } = true;
}

/// <summary>
/// Options controlling folder enumeration for <see cref="DocumentReader.ReadFolder(string, ReaderFolderOptions?, ReaderOptions?, System.Threading.CancellationToken)"/>.
/// </summary>
public sealed class ReaderFolderOptions {
    /// <summary>
    /// When true, enumerates all subdirectories. Default: true.
    /// </summary>
    public bool Recurse { get; set; } = true;

    /// <summary>
    /// Maximum number of files enumerated. Default: 500.
    /// </summary>
    public int MaxFiles { get; set; } = 500;

    /// <summary>
    /// Optional maximum total bytes across all enumerated files (best-effort).
    /// When null, no cap is enforced.
    /// </summary>
    public long? MaxTotalBytes { get; set; }

    /// <summary>
    /// Optional allowed extensions (lower/upper insensitive). When null, a default set is used.
    /// Examples: ".docx", ".xlsx", ".pptx", ".md".
    /// </summary>
    public IReadOnlyList<string>? Extensions { get; set; }

    /// <summary>
    /// When true, directory traversal skips reparse points (junctions/symlinks). Default: true.
    /// </summary>
    public bool SkipReparsePoints { get; set; } = true;

    /// <summary>
    /// When true, folder traversal is deterministic (ordinal path ordering). Default: true.
    /// </summary>
    public bool DeterministicOrder { get; set; } = true;
}

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
