#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how a <see cref="CsvDocument"/> is exposed as a forward-only data reader.
/// </summary>
public sealed class CsvDataReaderOptions
{
    /// <summary>
    /// Gets or sets an explicit schema used for typed reader columns.
    /// </summary>
    public CsvSchema? Schema { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether schema should be inferred before creating reader columns.
    /// </summary>
    public bool InferSchema { get; set; }

    /// <summary>
    /// Gets or sets the maximum row count inspected when <see cref="InferSchema"/> is enabled.
    /// </summary>
    public int SchemaSampleSize { get; set; } = 1000;
}
