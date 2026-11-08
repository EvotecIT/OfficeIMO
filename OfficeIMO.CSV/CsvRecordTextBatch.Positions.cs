#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal sealed partial class CsvRecordTextBatch
{
    private readonly (int? Start, int? End)[]? _positions;
    internal bool CapturesPositions => _positions is not null;
    internal int? CurrentPhysicalLineNumber => _positions is not null && (uint)_row < (uint)Count
        ? _positions[_row].Start : null;
    internal int? CurrentPhysicalEndLineNumber => _positions is not null && (uint)_row < (uint)Count
        ? _positions[_row].End : null;

    /// <summary>Keeps a detached row's original source location available to its reader.</summary>
    internal void CapturePosition(int? start, int? end)
    {
        if (_positions is not null) _positions[Count - 1] = (start, end);
    }
}
#endif
