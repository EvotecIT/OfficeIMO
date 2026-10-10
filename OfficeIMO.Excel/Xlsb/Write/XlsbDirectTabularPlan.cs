using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Captures a bounded native cell snapshot before writing a ZIP destination.
    /// The snapshot prevents a second read of mutable input during emission and
    /// lets the worksheet stream without retaining its complete BIFF12 payload.
    /// </summary>
    internal sealed class XlsbDirectTabularPlan : IDisposable {
        private const int MaximumCellSlots = 1_048_576;
        internal const byte DateKind = 5;
        private readonly int _cellSlots;
        private byte[] _kinds = Array.Empty<byte>();
        private ulong[] _payloads = Array.Empty<ulong>();
        private string?[]? _texts;

        private XlsbDirectTabularPlan(int rowCount, int columnCount, bool sharedStrings) {
            RowCount = rowCount;
            ColumnCount = columnCount;
            UsesSharedStrings = sharedStrings;
            _cellSlots = checked(rowCount * columnCount);
        }

        internal int RowCount { get; }
        internal int ColumnCount { get; }
        internal bool UsesSharedStrings { get; }
        internal uint DateStyle { get; private set; }
        internal byte KindAt(int slot) => _kinds[slot];
        internal ulong PayloadAt(int slot) => _payloads[slot];
        internal string TextAt(int slot) => _texts![slot]!;

        /// <summary>Large tables retain the existing staged worksheet route.</summary>
        internal static bool CanCapture(ExcelDirectTabularSource source) =>
            ((long)source.Rows.RowCount + (source.IncludeHeaders ? 1 : 0)) * source.Rows.ColumnCount <= MaximumCellSlots;

        internal static bool TryCreate(
            ExcelDocument document,
            ExcelSheet sheet,
            ExcelDirectTabularSource source,
            XlsbSharedStringTable? sharedStrings,
            CancellationToken cancellationToken,
            out XlsbDirectTabularPlan plan) {
            IExcelSheetTabularRowSource rows = source.Rows;
            int rowOffset = source.IncludeHeaders ? 1 : 0;
            int totalRows = checked(rows.RowCount + rowOffset);
            if (totalRows > 1_048_576 || rows.ColumnCount > 16_384) {
                throw new NotSupportedException("Native XLSB saving supports 1,048,576 rows and 16,384 columns per worksheet.");
            }
            if (!CanCapture(source)) throw new ArgumentOutOfRangeException(nameof(source));

            var candidate = new XlsbDirectTabularPlan(totalRows, rows.ColumnCount, sharedStrings != null);
            bool captured = false;
            try {
                if (candidate._cellSlots != 0) {
                    candidate._kinds = ArrayPool<byte>.Shared.Rent(candidate._cellSlots);
                    candidate._payloads = ArrayPool<ulong>.Shared.Rent(candidate._cellSlots);
                }
                if (source.IncludeHeaders) {
                    for (int column = 0; column < rows.ColumnCount; column++) {
                        candidate.CaptureText(column, rows.GetColumnName(column), sharedStrings);
                    }
                }
                object?[]? flatValues = rows.TryGetFlatValues(out object?[] values, out int flatColumnCount)
                    && flatColumnCount == rows.ColumnCount
                    && values.Length == checked(rows.RowCount * rows.ColumnCount)
                        ? values
                        : null;
                uint? dateStyle = null;
                for (int row = 0; row < rows.RowCount; row++) {
                    if ((row & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    object?[]? bufferedRow = flatValues == null
                        && rows.TryGetBufferedRow(row, out object?[]? valuesRow)
                        && valuesRow?.Length == rows.ColumnCount
                            ? valuesRow
                            : null;
                    int rowStart = checked((row + rowOffset) * rows.ColumnCount);
                    for (int column = 0; column < rows.ColumnCount; column++) {
                        int slot = rowStart + column;
                        object? rawValue = flatValues != null ? flatValues[row * rows.ColumnCount + column]
                            : bufferedRow != null ? bufferedRow[column] : rows.GetValue(row, column);
                        if (rawValue is DateTime date) {
                            double serial = ExcelDateSystemConverter.ToSerial(date, document.DateSystem);
                            dateStyle ??= sheet.GetOrCreateDirectTabularDateStyle(source.UseCellValueNumberFormats);
                            candidate._kinds[slot] = DateKind;
                            candidate._payloads[slot] = unchecked((ulong)BitConverter.DoubleToInt64Bits(serial));
                            continue;
                        }
                        ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(rawValue, source.PreserveMissingValues);
                        if (value.Kind == ExcelDirectTabularValueKind.Unsupported) {
                            plan = null!;
                            return false;
                        }
                        candidate._kinds[slot] = (byte)value.Kind;
                        switch (value.Kind) {
                            case ExcelDirectTabularValueKind.Text:
                                candidate.CaptureText(slot, value.Text ?? string.Empty, sharedStrings);
                                break;
                            case ExcelDirectTabularValueKind.Number:
                                candidate._payloads[slot] = unchecked((ulong)BitConverter.DoubleToInt64Bits(value.Number));
                                break;
                            case ExcelDirectTabularValueKind.Boolean:
                                candidate._payloads[slot] = value.Boolean ? 1UL : 0UL;
                                break;
                        }
                    }
                }
                cancellationToken.ThrowIfCancellationRequested();
                candidate.DateStyle = dateStyle ?? 0U;
                plan = candidate;
                captured = true;
                return true;
            } finally {
                if (!captured) candidate.Dispose();
            }
        }

        private void CaptureText(int slot, string text, XlsbSharedStringTable? sharedStrings) {
            CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            _kinds[slot] = (byte)ExcelDirectTabularValueKind.Text;
            if (sharedStrings != null) {
                _payloads[slot] = checked((ulong)sharedStrings.GetOrAdd(text));
            } else {
                _texts ??= ArrayPool<string?>.Shared.Rent(_cellSlots);
                _texts[slot] = text;
            }
        }

        public void Dispose() {
            byte[] kinds = _kinds;
            ulong[] payloads = _payloads;
            string?[]? texts = _texts;
            _kinds = Array.Empty<byte>();
            _payloads = Array.Empty<ulong>();
            _texts = null;
            if (kinds.Length != 0) ArrayPool<byte>.Shared.Return(kinds, clearArray: true);
            if (payloads.Length != 0) ArrayPool<ulong>.Shared.Return(payloads, clearArray: true);
            if (texts != null) ArrayPool<string?>.Shared.Return(texts, clearArray: true);
        }
    }
}
