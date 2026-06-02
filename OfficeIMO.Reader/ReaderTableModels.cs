using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Minimal table model for ingestion (columns + rows).
/// </summary>
public sealed class ReaderTable {
    /// <summary>
    /// Optional title/label.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Optional source-specific table kind/classification.
    /// </summary>
    public string? Kind { get; set; }

    /// <summary>
    /// Optional source-specific call or correlation identifier.
    /// </summary>
    public string? CallId { get; set; }

    /// <summary>
    /// Optional short descriptive summary.
    /// </summary>
    public string? Summary { get; set; }

    /// <summary>
    /// Optional stable short hash derived from the source payload.
    /// </summary>
    public string? PayloadHash { get; set; }

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
/// Minimal visual fence model for ingestion (kind + original language + payload).
/// </summary>
public sealed class ReaderVisual {
    /// <summary>
    /// Normalized visual kind (for example: "mermaid", "chart", or "network").
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Original fenced-code language label from markdown (for example: "ix-chart").
    /// </summary>
    public string Language { get; set; } = string.Empty;

    /// <summary>
    /// Raw visual payload/content from inside the fenced block.
    /// </summary>
    public string Content { get; set; } = string.Empty;

    /// <summary>
    /// Optional stable short hash derived from <see cref="Content"/>.
    /// </summary>
    public string? PayloadHash { get; set; }
}
