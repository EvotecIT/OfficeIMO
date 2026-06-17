#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how CSV serialization treats values that spreadsheet applications may interpret as formulas.
/// </summary>
public enum CsvFormulaInjectionPolicy
{
    /// <summary>
    /// Preserve values exactly as supplied.
    /// </summary>
    Preserve = 0,

    /// <summary>
    /// Prefix formula-like values with an apostrophe before writing them.
    /// </summary>
    Escape = 1
}
