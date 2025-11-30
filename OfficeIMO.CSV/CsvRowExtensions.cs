#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Convenience extension methods for <see cref="CsvRow"/>.
/// </summary>
public static class CsvRowExtensions
{
    /// <summary>
    /// Returns the field value as a string.
    /// </summary>
    public static string? AsString(this CsvRow row, string columnName) => row.Get<string>(columnName);

    /// <summary>
    /// Returns the field value as a 32-bit integer.
    /// </summary>
    public static int AsInt32(this CsvRow row, string columnName) => row.Get<int>(columnName);

    /// <summary>
    /// Returns the field value as a boolean.
    /// </summary>
    public static bool AsBoolean(this CsvRow row, string columnName) => row.Get<bool>(columnName);

    /// <summary>
    /// Returns the field value as a <see cref="DateTime"/>.
    /// </summary>
    public static DateTime AsDateTime(this CsvRow row, string columnName) => row.Get<DateTime>(columnName);
}
