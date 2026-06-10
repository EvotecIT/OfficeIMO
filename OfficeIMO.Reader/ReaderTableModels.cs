using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Inferred content kind for a reader table column.
/// </summary>
public enum ReaderTableColumnKind {
    /// <summary>The column has no non-empty body cells.</summary>
    Empty,

    /// <summary>The column has non-empty body cells that all look numeric.</summary>
    Numeric,

    /// <summary>The column has non-empty body cells that all look non-numeric.</summary>
    Text,

    /// <summary>The column contains both numeric-looking and non-numeric body cells.</summary>
    Mixed
}

/// <summary>
/// Inferred column profile for reader table ingestion consumers.
/// </summary>
public sealed class ReaderTableColumnProfile {
    /// <summary>Zero-based column index aligned to <see cref="ReaderTable.Columns"/>.</summary>
    public int Index { get; set; }

    /// <summary>Column name aligned to <see cref="ReaderTable.Columns"/>.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Inferred content kind for the column.</summary>
    public ReaderTableColumnKind Kind { get; set; }

    /// <summary>Number of non-empty body cells inspected for this column.</summary>
    public int NonEmptyCellCount { get; set; }

    /// <summary>Number of non-empty body cells that looked numeric.</summary>
    public int NumericCellCount { get; set; }

    /// <summary>Confidence score between 0 and 1 for the inferred <see cref="Kind"/>.</summary>
    public double Confidence { get; set; }

    /// <summary>True when the column profile is numeric.</summary>
    public bool IsNumeric => Kind == ReaderTableColumnKind.Numeric;
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
    /// Optional source location for the extracted table.
    /// </summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>
    /// Column headers.
    /// </summary>
    public IReadOnlyList<string> Columns { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Inferred column profiles aligned to <see cref="Columns"/>.
    /// </summary>
    public IReadOnlyList<ReaderTableColumnProfile> ColumnProfiles { get; set; } = Array.Empty<ReaderTableColumnProfile>();

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

    /// <summary>
    /// Optional source location for the extracted visual payload.
    /// </summary>
    public ReaderLocation? Location { get; set; }
}
