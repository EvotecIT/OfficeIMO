#nullable enable

namespace OfficeIMO.CSV;

#if NET8_0_OR_GREATER
/// <summary>
/// Receives CSV fields as transient spans during a single-pass read.
/// </summary>
public interface ICsvFieldSpanVisitor
{
    /// <summary>
    /// Visits one parsed field. The span is valid only for the duration of the call.
    /// </summary>
    /// <param name="recordIndex">Zero-based emitted record index.</param>
    /// <param name="fieldIndex">Zero-based field index within the record.</param>
    /// <param name="value">The field value. Do not capture the span beyond this method.</param>
    void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value);

    /// <summary>
    /// Optionally visits an escaped quoted field without forcing the parser to compact doubled quotes into a scratch buffer.
    /// </summary>
    /// <param name="recordIndex">Zero-based emitted record index.</param>
    /// <param name="fieldIndex">Zero-based field index within the record.</param>
    /// <param name="escapedValue">The quoted field content without surrounding quotes, preserving doubled quote escape sequences.</param>
    /// <param name="unescapedLength">Length of the field after CSV quote unescaping.</param>
    /// <returns><see langword="true" /> when the visitor consumed the escaped field; otherwise the parser falls back to <see cref="VisitField" /> with an unescaped span.</returns>
    bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
    {
        return false;
    }

    /// <summary>
    /// Visits one parsed string field. Implement this to avoid copying fields that were already materialized while parsing quoted records.
    /// </summary>
    /// <param name="recordIndex">Zero-based emitted record index.</param>
    /// <param name="fieldIndex">Zero-based field index within the record.</param>
    /// <param name="value">The field value.</param>
    void VisitFieldValue(int recordIndex, int fieldIndex, string value)
    {
        VisitField(recordIndex, fieldIndex, value.AsSpan());
    }
}

internal readonly struct CsvFieldSpanActionVisitor : ICsvFieldSpanVisitor
{
    private readonly CsvFieldSpanAction _action;

    public CsvFieldSpanActionVisitor(CsvFieldSpanAction action)
    {
        _action = action ?? throw new ArgumentNullException(nameof(action));
    }

    public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
    {
        _action(recordIndex, fieldIndex, value);
    }
}
#endif
