using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel;

/// <summary>
/// A chunk of Excel content extracted for ingestion (RAG/search/summarization).
/// </summary>
public sealed class ExcelExtractChunk {
    /// <summary>
    /// Stable, ASCII-only identifier (producer-defined).
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Source location information for citations and debugging.
    /// </summary>
    public ExcelExtractLocation Location { get; set; } = new ExcelExtractLocation();

    /// <summary>
    /// Plain text representation of the chunk.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional Markdown representation (usually a small table preview).
    /// </summary>
    public string? Markdown { get; set; }

    /// <summary>
    /// Structured tables extracted for this chunk.
    /// </summary>
    public IReadOnlyList<ExcelExtractTable> Tables { get; set; } = Array.Empty<ExcelExtractTable>();

    /// <summary>
    /// Optional warnings about truncation or unsupported content.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Generic Excel extraction location metadata.
/// </summary>
public sealed class ExcelExtractLocation {
    /// <summary>
    /// Optional source path (for example file path) used for citations.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Sheet name.
    /// </summary>
    public string? Sheet { get; set; }

    /// <summary>
    /// A1 range descriptor.
    /// </summary>
    public string? A1Range { get; set; }

    /// <summary>
    /// Optional producer-defined block index (chunk order).
    /// </summary>
    public int? BlockIndex { get; set; }
}

/// <summary>
/// Minimal table model for ingestion (columns + rows of strings).
/// </summary>
public sealed class ExcelExtractTable {
    /// <summary>
    /// Optional table title/label.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Column headers.
    /// </summary>
    public IReadOnlyList<string> Columns { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Table rows, aligned with <see cref="Columns"/>.
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
/// Chunking options for Excel extraction.
/// </summary>
public sealed class ExcelExtractChunkingOptions {
    /// <summary>
    /// Maximum characters per emitted chunk (best-effort).
    /// </summary>
    public int MaxChars { get; set; } = 8_000;

    /// <summary>
    /// Maximum number of table rows to include in a single chunk (best-effort).
    /// </summary>
    public int MaxTableRows { get; set; } = 200;
}

