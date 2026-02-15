using System;
using System.Collections.Generic;

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
    Text
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
}

/// <summary>
/// Options controlling folder enumeration for <see cref="DocumentReader.ReadFolder"/>.
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
