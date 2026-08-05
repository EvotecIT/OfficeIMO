using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private bool ReadCurrentRowRecordsFast() {
            bool checkCancellation = _cancellationToken.CanBeCanceled;
            byte[] bytes = _records.Buffer;
            int position = _records.Position;
            int length = _records.Length;
            while (position < length) {
                int firstTypeByte = bytes[position++];
                int recordType = firstTypeByte & 0x7F;
                if ((firstTypeByte & 0x80) != 0) {
                    recordType |= (bytes[position++] & 0x7F) << 7;
                }

                int current = bytes[position++];
                int recordSize = current & 0x7F;
                if ((current & 0x80) != 0) {
                    current = bytes[position++];
                    recordSize |= (current & 0x7F) << 7;
                    if ((current & 0x80) != 0) {
                        current = bytes[position++];
                        recordSize |= (current & 0x7F) << 14;
                        if ((current & 0x80) != 0) {
                            current = bytes[position++];
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
                    int rowIndex = Unsafe.ReadUnaligned<int>(ref bytes[payloadOffset]);
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
            bool checkCancellation = _cancellationToken.CanBeCanceled;
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
