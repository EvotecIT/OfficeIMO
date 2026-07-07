#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how duplicate CSV header names are handled when parsing input.
/// </summary>
public enum CsvDuplicateHeaderBehavior
{
    /// <summary>Preserve duplicate header names exactly as they appear in the CSV source.</summary>
    Preserve,

    /// <summary>Rename duplicate header names by appending a numeric suffix such as <c>_2</c>.</summary>
    Rename,

    /// <summary>Throw a <see cref="CsvException"/> when a duplicate header name is encountered.</summary>
    Throw
}
