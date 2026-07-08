#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Determines how CSV parse errors are handled.
/// </summary>
public enum CsvParseErrorAction
{
    /// <summary>
    /// Throw the first parse error.
    /// </summary>
    Throw,

    /// <summary>
    /// Collect the parse error and skip the malformed record when recovery is possible.
    /// </summary>
    SkipRow
}
