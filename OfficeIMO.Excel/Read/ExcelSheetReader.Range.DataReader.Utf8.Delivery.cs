#nullable enable

using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            /// <summary>
            /// Indexed rows pass their read token directly to window selection. They
            /// do not run XmlReader callbacks, so no temporary XML cancellation state
            /// or finally restoration is needed. Existing rows survive a canceled refill.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool ReadIndexedCore(CancellationToken cancellationToken) {
                _ct.ThrowIfCancellationRequested();
                cancellationToken.ThrowIfCancellationRequested();
                if (_closed) return false;
                ThrowIfXmlTextBufferingFailed();
                if (TryReadIndexedLogicalRow(cancellationToken, out _)) return true;
                ReleaseCurrentValueReferences();
                _currentRow = null;
                return false;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryReadIndexedLogicalRow(CancellationToken cancellationToken, out object?[] row) {
                row = Array.Empty<object?>();
                if (_nextLogicalRow > _lastRow) return false;
                bool hasPhysicalRow = _utf8Source!.SelectRow(_nextLogicalRow, _ct, cancellationToken);
                ReleaseCurrentValueReferences();
                row = hasPhysicalRow ? _currentValues : _blankRow;
                _currentRow = row;
                _currentRowIsBlank = !hasPhysicalRow;
                _currentRowActive = hasPhysicalRow;
                _currentRowFinished = !hasPhysicalRow;
                _nextLogicalRow++;
                return true;
            }
        }
    }
}
