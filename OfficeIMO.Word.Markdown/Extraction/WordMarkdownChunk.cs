using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Markdown;

/// <summary>
/// A chunk of Word content extracted as Markdown for ingestion (RAG/search/summarization).
/// </summary>
public sealed class WordMarkdownChunk {
    /// <summary>
    /// Stable, ASCII-only identifier (producer-defined).
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Source location information for citations and debugging.
    /// </summary>
    public WordMarkdownLocation Location { get; set; } = new WordMarkdownLocation();

    /// <summary>
    /// Plain text representation of the chunk. For this extractor it is Markdown.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Markdown representation of the chunk.
    /// </summary>
    public string Markdown { get; set; } = string.Empty;

    /// <summary>
    /// Optional warnings about truncation or unsupported content.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Generic Word-to-Markdown extraction location metadata.
/// </summary>
public sealed class WordMarkdownLocation {
    /// <summary>
    /// Optional source path (for example file path) used for citations.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Optional producer-defined block index (document-order).
    /// </summary>
    public int? BlockIndex { get; set; }

    /// <summary>
    /// Optional heading path label (for example "H1 > H2").
    /// </summary>
    public string? HeadingPath { get; set; }
}

/// <summary>
/// Chunking options for Word-to-Markdown extraction.
/// </summary>
public sealed class WordMarkdownChunkingOptions {
    /// <summary>
    /// Maximum characters per emitted chunk (best-effort; chunk boundaries are block-aligned).
    /// </summary>
    public int MaxChars { get; set; } = 8_000;

    /// <summary>
    /// When true, emit a final footnotes block. Default: true.
    /// </summary>
    public bool IncludeFootnotes { get; set; } = true;
}

