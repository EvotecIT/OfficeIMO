#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Describes how a CSV document should be loaded and iterated.
/// </summary>
public enum CsvLoadMode
{
    /// <summary>
    /// Loads all rows into memory. Transformations are allowed.
    /// </summary>
    InMemory,

    /// <summary>
    /// Keeps the data on disk and enumerates rows lazily. Transformations that require full materialization are not allowed.
    /// </summary>
    Stream
}
