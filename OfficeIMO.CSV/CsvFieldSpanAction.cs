#nullable enable

namespace OfficeIMO.CSV;

#if NET8_0_OR_GREATER
/// <summary>
/// Receives a parsed CSV field as a transient character span.
/// </summary>
/// <param name="recordIndex">Zero-based emitted record index after skipped records and comments.</param>
/// <param name="fieldIndex">Zero-based field index inside the current record.</param>
/// <param name="value">Field value. The span is only valid for the duration of the callback.</param>
public delegate void CsvFieldSpanAction(int recordIndex, int fieldIndex, ReadOnlySpan<char> value);
#endif
