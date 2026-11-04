using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    internal sealed partial class XlsbDirectTabularPlan {
        private static bool TryCreateBuffered(
            ExcelDateSystem dateSystem,
            ExcelSheet sheet,
            ExcelDirectTabularSource source,
            ExcelCellValueBuffer cells,
            XlsbSharedStringTable? sharedStrings,
            CancellationToken cancellationToken,
            out XlsbDirectTabularPlan plan) {
            var candidate = CreateCandidate(source, sharedStrings);
            bool captured = false;
            try {
                candidate.InitializeStorage(source, sharedStrings, cancellationToken);
                if (!candidate.TryCaptureBufferedCells(dateSystem, sheet, source, cells,
                        sharedStrings, cancellationToken)) {
                    plan = null!;
                    return false;
                }
                plan = candidate;
                captured = true;
                return true;
            } finally {
                if (!captured) candidate.Dispose();
            }
        }

        /// <summary>
        /// Captures the same pre-write snapshot directly from typed cells. Missing cells
        /// stay absent, while an explicit blank retains the CellValue empty-text contract.
        /// No source buffer is borrowed during destination writes.
        /// </summary>
        private bool TryCaptureBufferedCells(
            ExcelDateSystem dateSystem,
            ExcelSheet sheet,
            ExcelDirectTabularSource source,
            ExcelCellValueBuffer cells,
            XlsbSharedStringTable? sharedStrings,
            CancellationToken cancellationToken) {
            _kinds.AsSpan(0, _cellSlots).Clear();
            uint? dateStyle = null;
            int visited = 0;
            for (int row = 0; row < cells.WrittenRowCount; row++) {
                cancellationToken.ThrowIfCancellationRequested();
                int rowStart = checked((cells.GetRowNumber(row) - 1) * ColumnCount);
                int end = cells.GetEndCellIndex(row);
                for (int index = cells.GetFirstCellIndex(row); index < end; index++) {
                    if ((visited++ & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    int slot = checked(rowStart + cells.GetColumn(index) - 1);
                    ExcelBufferedCell cell = cells.GetCell(index);
                    switch (cell.Kind) {
                        case ExcelBufferedCellKind.Blank:
                            CaptureText(slot, string.Empty, sharedStrings);
                            break;
                        case ExcelBufferedCellKind.Text:
                            CaptureText(slot, (string)cell.Reference!, sharedStrings);
                            break;
                        case ExcelBufferedCellKind.Number:
                            double number = cell.NumberValue;
                            if (double.IsNaN(number) || double.IsInfinity(number)) return false;
                            _kinds[slot] = (byte)ExcelDirectTabularValueKind.Number;
                            _payloads[slot] = unchecked((ulong)cell.Bits);
                            break;
                        case ExcelBufferedCellKind.Int32:
                        case ExcelBufferedCellKind.Int64:
                            _kinds[slot] = (byte)ExcelDirectTabularValueKind.Number;
                            _payloads[slot] = unchecked((ulong)BitConverter.DoubleToInt64Bits((double)cell.Bits));
                            break;
                        case ExcelBufferedCellKind.DateTime:
                            double serial = ExcelDateSystemConverter.ToSerial(cell.DateTimeValue, dateSystem);
                            dateStyle ??= sheet.GetOrCreateDirectTabularDateStyle(source.UseCellValueNumberFormats);
                            _kinds[slot] = DateKind;
                            _payloads[slot] = unchecked((ulong)BitConverter.DoubleToInt64Bits(serial));
                            break;
                        case ExcelBufferedCellKind.Boolean:
                            _kinds[slot] = (byte)ExcelDirectTabularValueKind.Boolean;
                            _payloads[slot] = cell.Bits != 0L ? 1UL : 0UL;
                            break;
                        case ExcelBufferedCellKind.Value:
                            ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(cell.Reference, preserveMissingValues: true);
                            switch (value.Kind) {
                                case ExcelDirectTabularValueKind.Unsupported:
                                    return false;
                                case ExcelDirectTabularValueKind.Text:
                                    CaptureText(slot, value.Text ?? string.Empty, sharedStrings);
                                    break;
                                case ExcelDirectTabularValueKind.Number:
                                    _kinds[slot] = (byte)value.Kind;
                                    _payloads[slot] = unchecked((ulong)BitConverter.DoubleToInt64Bits(value.Number));
                                    break;
                                case ExcelDirectTabularValueKind.Boolean:
                                    _kinds[slot] = (byte)value.Kind;
                                    _payloads[slot] = value.Boolean ? 1UL : 0UL;
                                    break;
                            }
                            break;
                    }
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            DateStyle = dateStyle ?? 0U;
            return true;
        }
    }
}
