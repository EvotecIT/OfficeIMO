#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Describes CSV read progress reported while records are emitted.
/// </summary>
public readonly struct CsvProgress
{
    /// <summary>
    /// Initializes a new progress snapshot.
    /// </summary>
    public CsvProgress(long recordsRead, int lineNumber)
    {
        RecordsRead = recordsRead;
        LineNumber = lineNumber;
    }

    /// <summary>
    /// Gets the number of emitted records.
    /// </summary>
    public long RecordsRead { get; }

    /// <summary>
    /// Gets the current physical line number.
    /// </summary>
    public int LineNumber { get; }
}
