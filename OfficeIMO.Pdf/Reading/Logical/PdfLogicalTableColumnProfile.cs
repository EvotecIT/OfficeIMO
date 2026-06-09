namespace OfficeIMO.Pdf;

/// <summary>
/// Describes the inferred content kind for a normalized logical PDF table column.
/// </summary>
public enum PdfLogicalTableColumnKind {
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
/// Inferred column profile for adapters that need stable table typing hints without re-analyzing cell text.
/// </summary>
public sealed class PdfLogicalTableColumnProfile {
    internal PdfLogicalTableColumnProfile(
        int index,
        string name,
        PdfLogicalTableColumnKind kind,
        int nonEmptyCellCount,
        int numericCellCount,
        double confidence) {
        Index = index;
        Name = name ?? string.Empty;
        Kind = kind;
        NonEmptyCellCount = nonEmptyCellCount;
        NumericCellCount = numericCellCount;
        Confidence = confidence;
    }

    /// <summary>Zero-based column index in the normalized table.</summary>
    public int Index { get; }

    /// <summary>Inferred column name aligned to <see cref="PdfLogicalTableData.Columns"/>.</summary>
    public string Name { get; }

    /// <summary>Inferred content kind for non-empty body cells.</summary>
    public PdfLogicalTableColumnKind Kind { get; }

    /// <summary>Number of non-empty body cells inspected for this column.</summary>
    public int NonEmptyCellCount { get; }

    /// <summary>Number of non-empty body cells that looked numeric.</summary>
    public int NumericCellCount { get; }

    /// <summary>
    /// Confidence score between 0 and 1 for the inferred <see cref="Kind"/> based on inspected body cells.
    /// </summary>
    public double Confidence { get; }

    /// <summary>True when the column is confidently numeric.</summary>
    public bool IsNumeric => Kind == PdfLogicalTableColumnKind.Numeric;
}
