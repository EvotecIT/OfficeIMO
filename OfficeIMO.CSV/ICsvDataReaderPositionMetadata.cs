#nullable enable

namespace OfficeIMO.CSV;

/// <summary>Exposes the current logical record and optional physical source-line position.</summary>
public interface ICsvDataReaderPositionMetadata
{
    /// <summary>
    /// Gets the one-based data-record number after header, comment, blank, and skipped-record handling.
    /// Returns zero when the reader is not positioned on a row.
    /// </summary>
    long RecordNumber { get; }

    /// <summary>
    /// Gets the one-based physical source line on which the current record starts, when retained by the reader path.
    /// Returns <c>null</c> when unavailable or when the reader is not positioned on a row.
    /// </summary>
    int? PhysicalLineNumber { get; }

    /// <summary>
    /// Gets the one-based physical source line on which the current record ends, when retained by the reader path.
    /// This differs from <see cref="PhysicalLineNumber"/> for multiline quoted records.
    /// </summary>
    int? PhysicalEndLineNumber { get; }
}

internal interface ICsvDataReaderPositionSource
{
    int? CurrentPhysicalLineNumber { get; }

    int? CurrentPhysicalEndLineNumber { get; }
}
