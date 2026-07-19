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
/// Optional table confidence and geometry diagnostics for reader ingestion consumers.
/// </summary>
public sealed class ReaderTableDiagnostics {
    /// <summary>Overall confidence score between 0 and 1 based on source-specific table signals.</summary>
    public double Confidence { get; set; }

    /// <summary>Confidence score between 0 and 1 for the inferred table schema.</summary>
    public double SchemaConfidence { get; set; }

    /// <summary>Ratio between 0 and 1 of non-empty cells to expected cells in the source table.</summary>
    public double CellCompleteness { get; set; }

    /// <summary>Confidence score between 0 and 1 for source column geometry.</summary>
    public double ColumnGeometryConfidence { get; set; }

    /// <summary>Number of source rows detected before reader row truncation.</summary>
    public int SourceRowCount { get; set; }

    /// <summary>Expected source cell count from source rows multiplied by inferred columns.</summary>
    public int ExpectedCellCount { get; set; }

    /// <summary>Number of non-empty source cells detected.</summary>
    public int FilledCellCount { get; set; }

    /// <summary>Number of expected source cells that were empty or unavailable.</summary>
    public int MissingCellCount { get; set; }

    /// <summary>Left edge of the source table geometry in source units, when available.</summary>
    public double XStart { get; set; }

    /// <summary>Right edge of the source table geometry in source units, when available.</summary>
    public double XEnd { get; set; }

    /// <summary>Top coordinate of the source table geometry in source units, when available.</summary>
    public double YTop { get; set; }

    /// <summary>Bottom coordinate of the source table geometry in source units, when available.</summary>
    public double YBottom { get; set; }

    /// <summary>Source table width in source units.</summary>
    public double Width { get; set; }

    /// <summary>Source table height in source units.</summary>
    public double Height { get; set; }

    /// <summary>True when source table geometry was available.</summary>
    public bool HasGeometry { get; set; }
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
    /// Optional source-specific table confidence and geometry diagnostics.
    /// </summary>
    public ReaderTableDiagnostics? Diagnostics { get; set; }

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
    /// Optional source location for the extracted visual.
    /// </summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>
    /// Optional source-specific visual name, for example a PDF image resource name.
    /// </summary>
    public string? SourceName { get; set; }

    /// <summary>
    /// Optional MIME type when the visual source exposes one.
    /// </summary>
    public string? MimeType { get; set; }

    /// <summary>
    /// Intrinsic visual width in source units or pixels, when available.
    /// </summary>
    public double? Width { get; set; }

    /// <summary>
    /// Intrinsic visual height in source units or pixels, when available.
    /// </summary>
    public double? Height { get; set; }

    /// <summary>
    /// Left or X coordinate of the placed visual in source units, when available.
    /// </summary>
    public double? X { get; set; }

    /// <summary>
    /// Bottom or Y coordinate of the placed visual in source units, when available.
    /// </summary>
    public double? Y { get; set; }

    /// <summary>
    /// Placed visual width in source units, when available.
    /// </summary>
    public double? PlacedWidth { get; set; }

    /// <summary>
    /// Placed visual height in source units, when available.
    /// </summary>
    public double? PlacedHeight { get; set; }

    /// <summary>
    /// Number of detected placement instances represented by this visual.
    /// </summary>
    public int PlacementCount { get; set; }

    /// <summary>
    /// True when source placement geometry was available.
    /// </summary>
    public bool HasGeometry { get; set; }

    /// <summary>
    /// True when source placement geometry was axis-aligned, when known.
    /// </summary>
    public bool? IsAxisAligned { get; set; }
}
