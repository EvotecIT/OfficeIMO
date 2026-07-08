#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how a <see cref="CsvDocument"/> is projected into a <see cref="System.Data.DataTable"/>.
/// </summary>
public sealed class CsvDataTableOptions
{
    /// <summary>
    /// Gets or sets the DataTable name. Defaults to CsvData.
    /// </summary>
    public string? TableName { get; set; }

    /// <summary>
    /// Gets or sets an explicit schema used for typed DataTable columns.
    /// </summary>
    public CsvSchema? Schema { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether schema should be inferred before creating columns.
    /// </summary>
    public bool InferSchema { get; set; }

    /// <summary>
    /// Gets or sets the maximum row count inspected when <see cref="InferSchema"/> is enabled.
    /// </summary>
    public int SchemaSampleSize { get; set; } = 1000;
}
