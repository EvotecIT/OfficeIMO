using System.Collections.Generic;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// A chunk of PowerPoint content extracted as Markdown for ingestion (RAG/search/summarization).
/// </summary>
public sealed class PowerPointExtractChunk {
    /// <summary>
    /// Stable, ASCII-only identifier (producer-defined).
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Source location information for citations and debugging.
    /// </summary>
    public PowerPointExtractLocation Location { get; set; } = new PowerPointExtractLocation();

    /// <summary>
    /// Plain text representation of the chunk.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional Markdown representation of the chunk.
    /// </summary>
    public string? Markdown { get; set; }

    /// <summary>
    /// Optional warnings about truncation or unsupported content.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Generic PowerPoint extraction location metadata.
/// </summary>
public sealed class PowerPointExtractLocation {
    /// <summary>
    /// Optional source path (for example file path) used for citations.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// 1-based slide number.
    /// </summary>
    public int? Slide { get; set; }

    /// <summary>
    /// Optional producer-defined block index (chunk order).
    /// </summary>
    public int? BlockIndex { get; set; }
}

/// <summary>
/// Chunking options for PowerPoint extraction.
/// </summary>
public sealed class PowerPointExtractChunkingOptions {
    /// <summary>
    /// Maximum characters per emitted chunk (best-effort).
    /// </summary>
    public int MaxChars { get; set; } = 8_000;
}

