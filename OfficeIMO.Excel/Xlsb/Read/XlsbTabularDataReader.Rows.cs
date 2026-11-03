using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private bool ReadCurrentRowRecordsFast() {
            bool checkCancellation = CanCancelCurrentRead;
            byte[] bytes = _records.Buffer;
            // Discovery validated every record boundary in this immutable buffer.
            ref byte data = ref MemoryMarshal.GetReference(bytes.AsSpan());
            int position = _records.Position;
            int length = _records.Length;
            while (position < length) {
                int firstTypeByte = Unsafe.Add(ref data, position++);
                int recordType = firstTypeByte & 0x7F;
                if ((firstTypeByte & 0x80) != 0) {
                    recordType |= (Unsafe.Add(ref data, position++) & 0x7F) << 7;
                }

                int current = Unsafe.Add(ref data, position++);
                int recordSize = current & 0x7F;
                if ((current & 0x80) != 0) {
                    current = Unsafe.Add(ref data, position++);
                    recordSize |= (current & 0x7F) << 7;
                    if ((current & 0x80) != 0) {
                        current = Unsafe.Add(ref data, position++);
                        recordSize |= (current & 0x7F) << 14;
                        if ((current & 0x80) != 0) {
                            current = Unsafe.Add(ref data, position++);
                            recordSize |= (current & 0x7F) << 21;
                        }
                    }
                }

                int payloadOffset = position;
                position += recordSize;
                if (checkCancellation) {
                    CheckCancellation();
                }
                if (recordType == BrtRowHdr) {
                    int rowIndex = Unsafe.ReadUnaligned<int>(ref Unsafe.Add(ref data, payloadOffset));
                    if (rowIndex <= _lastDataRow) {
                        _pendingRowIndex = rowIndex;
                        _hasPendingRow = true;
                        _records.Position = position;
                        return true;
                    }
                    continue;
                }
                if (recordType == BrtEndSheetData) {
                    _reachedEndSheetData = true;
                    _records.Position = position;
                    return true;
                }
                switch (recordType) {
                    case BrtCellRk:
                        StoreValidatedRkCell(bytes, payloadOffset);
                        break;
                    case BrtCellReal:
                        StoreValidatedRealCell(bytes, payloadOffset);
                        break;
                    case BrtCellIsst:
                        StoreValidatedSharedStringCell(bytes, payloadOffset);
                        break;
                    default:
                        if (IsCellRecord(recordType)) {
                            StoreCellFast(bytes, recordType, payloadOffset);
                        }
                        break;
                }
            }

            _records.Position = position;
            return false;
        }

        private bool ReadCurrentRowRecordsConverted() {
            bool checkCancellation = CanCancelCurrentRead;
            while (_records.TryReadValidated(out XlsbRecordSlice record)) {
                if (checkCancellation) {
                    CheckCancellation();
                }
                if (record.Type == BrtRowHdr) {
                    if (TrySetPendingRow(record)) {
                        return true;
                    }
                    continue;
                }
                if (record.Type == BrtEndSheetData) {
                    _reachedEndSheetData = true;
                    return true;
                }
                if (IsCellRecord(record.Type)) {
                    StoreCell(record);
                }
            }
            return false;
        }
    }
}
