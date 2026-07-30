using OfficeIMO.Excel.Xlsb.Biff12;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private static void DiscoverDataColumns(
            Stream worksheetPart,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            CancellationToken cancellationToken,
            out int firstColumn,
            out int lastColumn,
            out int firstDataRow,
            out int lastDataRow) {
            if (!worksheetPart.CanSeek) {
                throw new InvalidOperationException(
                    "XLSB reads require a seekable worksheet part for schema discovery.");
            }

            firstColumn = int.MaxValue;
            lastColumn = -1;
            firstDataRow = -1;
            lastDataRow = -1;
            int currentRow = -1;
            long startPosition = worksheetPart.Position;
            try {
                using var scanner = new XlsbStreamRecordSliceReader(
                    worksheetPart,
                    limits.MaxRecordBytes,
                    recordBudget,
                    leaveOpen: true);
                bool inSheetData = false;
                int recordsSinceCancellationCheck = 0;
                cancellationToken.ThrowIfCancellationRequested();
                while (scanner.TryRead(out XlsbRecordSlice record)) {
                    recordsSinceCancellationCheck++;
                    if ((recordsSinceCancellationCheck & 1023) == 0) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    if (record.Type == BrtBeginSheetData) {
                        inSheetData = true;
                        continue;
                    }
                    if (!inSheetData) {
                        continue;
                    }
                    if (record.Type == BrtEndSheetData) {
                        break;
                    }
                    if (record.Type == BrtRowHdr) {
                        currentRow = ValidateRowHeader(record);
                        continue;
                    }
                    if (!IsCellRecord(record.Type)) {
                        continue;
                    }

                    if (firstDataRow < 0 && currentRow >= 0) {
                        firstDataRow = currentRow;
                    }
                    if (currentRow >= 0) {
                        lastDataRow = currentRow;
                    }
                    int column = record.CreateCursor().ReadInt32();
                    if (column < 0 || column >= A1.MaxColumns) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
                    }

                    firstColumn = Math.Min(firstColumn, column);
                    lastColumn = Math.Max(lastColumn, column);
                }
            } finally {
                worksheetPart.Position = startPosition;
            }
        }
    }
}
