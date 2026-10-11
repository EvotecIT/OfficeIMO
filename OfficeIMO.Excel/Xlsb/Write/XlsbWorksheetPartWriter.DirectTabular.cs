using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    internal static partial class XlsbWorksheetPartWriter {
        /// <summary>Emits each source value once to an owned worksheet entry before package publication.</summary>
        internal static bool TryWriteDirectTabular(
            ExcelDocument document,
            ExcelSheet sheet,
            ExcelDirectTabularSource source,
            Stream output,
            CancellationToken cancellationToken,
            XlsbSharedStringTable? sharedStrings = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));

            IExcelSheetTabularRowSource rows = source.Rows;
            int rowOffset = source.IncludeHeaders ? 1 : 0;
            int totalRows = ValidateDirectTabularDimensions(source);
            object?[]? flatValues = rows.TryGetFlatValues(out object?[] candidateValues, out int flatColumnCount)
                && flatColumnCount == rows.ColumnCount
                && candidateValues.Length == checked(rows.RowCount * rows.ColumnCount)
                    ? candidateValues
                    : null;

            using var writer = new XlsbDirectRecordWriter(output);
            writer.WriteRecord(129); // BrtBeginSheet
            writer.WriteHeader(BrtWsDim, 16);
            WriteDirectDimension(writer, totalRows, rows.ColumnCount);
            writer.WriteRecord(BrtBeginSheetData);

            if (source.IncludeHeaders && rows.ColumnCount != 0) {
                WriteDirectRowHeader(writer, 0, rows.ColumnCount);
                for (int column = 0; column < rows.ColumnCount; column++) {
                    WriteDirectTextCell(writer, column, rows.GetColumnName(column), sharedStrings);
                }
            }

            bool canCancel = cancellationToken.CanBeCanceled;
            uint? dateStyle = null;
            for (int row = 0; row < rows.RowCount; row++) {
                if (canCancel && (row & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                int zeroBasedRow = row + rowOffset;
                if (rows.ColumnCount != 0) {
                    WriteDirectRowHeader(writer, zeroBasedRow, rows.ColumnCount);
                }
                object?[]? bufferedRow = flatValues == null
                    && rows.TryGetBufferedRow(row, out object?[]? candidateRow)
                    && candidateRow?.Length == rows.ColumnCount
                        ? candidateRow
                        : null;
                bool wroteCell = false;
                for (int column = 0; column < rows.ColumnCount; column++) {
                    object? rawValue = flatValues != null
                        ? flatValues[checked((row * rows.ColumnCount) + column)]
                        : bufferedRow != null
                            ? bufferedRow[column]
                            : rows.GetValue(row, column);
                    if (rawValue is DateTime date) {
                        double serial = ExcelDateSystemConverter.ToSerial(date, document.DateSystem);
                        dateStyle ??= sheet.GetOrCreateDirectTabularDateStyle(source.UseCellValueNumberFormats);
                        writer.WriteNumberCell(BrtCellReal, column, serial, dateStyle.Value);
                        wroteCell = true;
                        continue;
                    }
                    ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(rawValue, source.PreserveMissingValues);
                    if (value.Kind == ExcelDirectTabularValueKind.Empty) continue;
                    wroteCell = true;
                    switch (value.Kind) {
                        case ExcelDirectTabularValueKind.Text:
                            WriteDirectTextCell(writer, column, value.Text ?? string.Empty, sharedStrings);
                            break;
                        case ExcelDirectTabularValueKind.Number:
                            WriteDirectNumberCell(writer, column, value.Number);
                            break;
                        case ExcelDirectTabularValueKind.Boolean:
                            WriteDirectBooleanCell(writer, column, value.Boolean);
                            break;
                        default:
                            return false;
                    }
                }
                if (!wroteCell && rows.ColumnCount != 0) {
                    // Retain an explicitly imported empty row without storing text
                    // or expanding every missing field into a cell record.
                    writer.WriteHeader(BrtCellBlank, 8);
                    writer.WriteUInt32(checked((uint)(rows.ColumnCount - 1)));
                    writer.WriteUInt32(0U);
                }
            }

            writer.WriteRecord(BrtEndSheetData);
            writer.WriteRecord(BrtEndSheet);
            writer.Flush();
            return true;
        }

        /// <summary>Retains the staged worksheet's bounded initial reservation for an owned package.</summary>
        internal static int EstimateDirectWorksheetCapacity(ExcelDirectTabularSource source) =>
            EstimateDirectWorksheetCapacity(ValidateDirectTabularDimensions(source), source.Rows.ColumnCount);

        private static int ValidateDirectTabularDimensions(ExcelDirectTabularSource source) {
            int totalRows = checked(source.Rows.RowCount + (source.IncludeHeaders ? 1 : 0));
            if (totalRows > 1_048_576 || source.Rows.ColumnCount > 16_384) {
                throw new NotSupportedException("Native XLSB saving supports 1,048,576 rows and 16,384 columns per worksheet.");
            }
            return totalRows;
        }

        private static void WriteDirectDimension(XlsbDirectRecordWriter writer, int rowCount, int columnCount) {
            int lastRow = Math.Max(1, rowCount);
            int lastColumn = Math.Max(1, columnCount);
            writer.WriteUInt32(0U);
            writer.WriteUInt32(checked((uint)(lastRow - 1)));
            writer.WriteUInt32(0U);
            writer.WriteUInt32(checked((uint)(lastColumn - 1)));
        }

        private static void WriteDirectRowHeader(
            XlsbDirectRecordWriter writer,
            int zeroBasedRow,
            int columnCount) {
            writer.WriteRowHeader(BrtRowHdr, zeroBasedRow, columnCount, DefaultRowProperties);
        }

        private static void WriteDirectTextCell(
            XlsbDirectRecordWriter writer,
            int zeroBasedColumn,
            string value,
            XlsbSharedStringTable? sharedStrings) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            if (sharedStrings == null) {
                writer.WriteTextCell(BrtCellSt, zeroBasedColumn, value);
            } else {
                writer.WriteSharedStringCell(BrtCellIsst, zeroBasedColumn, sharedStrings.GetOrAdd(value));
            }
        }

        private static void WriteDirectNumberCell(XlsbDirectRecordWriter writer, int zeroBasedColumn, double value) {
            writer.WriteNumberCell(BrtCellReal, zeroBasedColumn, value);
        }

        private static void WriteDirectBooleanCell(XlsbDirectRecordWriter writer, int zeroBasedColumn, bool value) {
            writer.WriteBooleanCell(BrtCellBool, zeroBasedColumn, value);
        }

        private static int EstimateDirectWorksheetCapacity(int rowCount, int columnCount) {
            const int maximumInitialCapacity = 16 * 1024 * 1024;
            int spanCount = checked((columnCount + 1023) / 1024);
            // Bound the dense-cell assumption so a very wide, sparse table does not
            // reserve a large buffer before the actual values establish its density.
            int estimatedDenseColumns = Math.Min(columnCount, 128);
            long bytesPerRow = 19L + spanCount * 8L + estimatedDenseColumns * 24L;
            long estimate = 64L + rowCount * bytesPerRow;
            return (int)Math.Max(256L, Math.Min(maximumInitialCapacity, estimate));
        }

    }
}
