using OfficeIMO.Markdown;
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
    /// Optional structured visual fence metadata extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderVisual>? Visuals { get; set; }

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
    /// Optional 1-based source start line number when the reader can map it accurately.
    /// </summary>
    public int? StartLine { get; set; }

    /// <summary>
    /// Optional 1-based source end line number when the reader can map it accurately.
    /// </summary>
    public int? EndLine { get; set; }

    /// <summary>
    /// Optional 1-based normalized markdown start line used for parser-aware markdown provenance.
    /// </summary>
    public int? NormalizedStartLine { get; set; }

    /// <summary>
    /// Optional 1-based normalized markdown end line used for parser-aware markdown provenance.
    /// </summary>
    public int? NormalizedEndLine { get; set; }

    /// <summary>
    /// Optional heading path label (for example "H1 &gt; H2").
    /// </summary>
    public string? HeadingPath { get; set; }

    /// <summary>
    /// Optional unique slug/anchor for the active heading section.
    /// </summary>
    public string? HeadingSlug { get; set; }

    /// <summary>
    /// Optional source block kind label (for example "heading", "paragraph", "code").
    /// </summary>
    public string? SourceBlockKind { get; set; }

    /// <summary>
    /// Optional deterministic anchor for the first logical block included in the chunk.
    /// For markdown this can identify a heading section or a sub-block within that section.
    /// </summary>
    public string? BlockAnchor { get; set; }

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
