#nullable enable

namespace OfficeIMO.CSV;

#if NET8_0_OR_GREATER
/// <summary>
/// Receives header-aware CSV data rows as transient field spans during a single-pass read.
/// </summary>
public interface ICsvRowFieldSpanVisitor
{
    /// <summary>
    /// Starts a data row. The header has already been discovered, normalized, and validated.
    /// </summary>
    /// <param name="header">Normalized header names for the row.</param>
    /// <param name="rowIndex">Zero-based emitted data row index.</param>
    void BeginRow(IReadOnlyList<string> header, int rowIndex);

    /// <summary>
    /// Visits one data field. The span is valid only for the duration of the call.
    /// </summary>
    /// <param name="rowIndex">Zero-based emitted data row index.</param>
    /// <param name="fieldIndex">Zero-based field index within the normalized header width.</param>
    /// <param name="value">The field value. Do not capture the span beyond this method.</param>
    void VisitField(int rowIndex, int fieldIndex, ReadOnlySpan<char> value);

    /// <summary>
    /// Visits one parsed string field. Implement this to avoid copying fields that were already materialized while parsing quoted records.
    /// </summary>
    /// <param name="rowIndex">Zero-based emitted data row index.</param>
    /// <param name="fieldIndex">Zero-based field index within the normalized header width.</param>
    /// <param name="value">The field value.</param>
    void VisitFieldValue(int rowIndex, int fieldIndex, string value)
    {
        VisitField(rowIndex, fieldIndex, value.AsSpan());
    }

    /// <summary>
    /// Completes a data row after all emitted field spans have been visited.
    /// </summary>
    /// <param name="rowIndex">Zero-based emitted data row index.</param>
    /// <param name="fieldCount">Number of fields parsed from the source row before column-count alignment.</param>
    void EndRow(int rowIndex, int fieldCount);
}
#endif
