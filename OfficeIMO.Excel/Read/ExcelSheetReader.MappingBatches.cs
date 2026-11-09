using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelSheetReader {
    internal static bool CanCaptureMappingBatch(DbDataReader reader) =>
        reader is ExcelXmlRangeDataReader { CanCaptureMappingBatch: true };

    internal static bool TryCaptureMappingBatch(DbDataReader reader, int count,
        CancellationToken ct, out DbDataReader? batch, out int captured) {
        batch = null;
        captured = 0;
        return reader is ExcelXmlRangeDataReader source && source.TryCaptureMappingBatch(count, ct, out batch, out captured);
    }

    private sealed partial class ExcelXmlRangeDataReader {
        internal bool CanCaptureMappingBatch => !_closed && !_trackCellPresence &&
            _utf8Source is { CanCopyMappingRows: true } &&
            !_owner._opt.NumericAsDecimal && _owner._opt.CellValueConverter is null;

        private ExcelXmlRangeDataReader(ExcelXmlRangeDataReader source,
            ExcelSheetReader owner, ExcelUtf8RangeRowSource rows, int count)
            : this(owner, 1, source._firstColumn, count, source._lastColumn,
                source._fieldCount, false, owner._opt, CancellationToken.None,
                rows, source._firstColumn, rowsAlreadyQualified: true) {
            _columnNames = source._columnNames;
            _columnTypes = source._columnTypes;
        }

        internal bool TryCaptureMappingBatch(int preferredBatchSize, CancellationToken ct,
            out DbDataReader? batch, out int count) {
            batch = null;
            count = 0;
            _ct.ThrowIfCancellationRequested();
            ct.ThrowIfCancellationRequested();
            if (!CanCaptureMappingBatch) return false;
            if (preferredBatchSize != 0 && _nextLogicalRow > _lastRow) return true;
            count = preferredBatchSize == 0 ? 0 : Math.Min(Math.Max(1, preferredBatchSize), _lastRow - _nextLogicalRow + 1);
            var owner = new ExcelSheetReader(_owner._sheetName, _owner._worksheetPartName,
                _owner._sst, _owner._styles, _owner._opt.Clone(), _owner._dateSystem,
                _owner._partBufferReader!, _owner._richValueErrors);
            if (!_utf8Source!.TryCopyMappingRows(owner, _nextLogicalRow, count,
                    _utf8SourceOrdinalOffset, _fieldCount, ct, out var rows)) return false;
            try { batch = new ExcelXmlRangeDataReader(this, owner, rows!, count); }
            catch { rows!.Dispose(); throw; }
            if (count != 0) {
                _nextLogicalRow += count;
                _utf8Source.AdvanceMappingRows(_nextLogicalRow);
                ReleaseCurrentValueReferences();
                _currentRow = null;
            }
            return true;
        }
    }
}
