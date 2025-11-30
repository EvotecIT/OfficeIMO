#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Represents a single schema validation error within a CSV document.
/// </summary>
public readonly record struct CsvValidationError(int RowIndex, string ColumnName, string Message);
