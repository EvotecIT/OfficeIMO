namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private bool ReadCurrentRowRecordsFast() {
            bool checkCancellation = _cancellationToken.CanBeCanceled;
            while (_records.TryRead(out XlsbRecordSlice record)) {
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
                    StoreCellFast(record);
                }
            }
            return false;
        }

        private bool ReadCurrentRowRecordsConverted() {
            bool checkCancellation = _cancellationToken.CanBeCanceled;
            while (_records.TryRead(out XlsbRecordSlice record)) {
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
