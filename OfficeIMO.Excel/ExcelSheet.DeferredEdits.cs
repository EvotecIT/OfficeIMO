using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>The changes a workbook that has not opened its SDK package records without it.</summary>
    internal enum ExcelDeferredCellKind : byte {
        /// <summary>Text that becomes a shared string, or cell text past the shared-string promotion limit.</summary>
        Text,
        /// <summary>A double.</summary>
        Number,
        /// <summary>An integer that fits a number cell.</summary>
        Integer,
        /// <summary>A decimal that fits a number cell, kept as its invariant text.</summary>
        Decimal,
        Boolean,
        /// <summary>A date, written as its serial with the automatic date format.</summary>
        Date,
        /// <summary>A style change; the cell keeps its value.</summary>
        Style,
    }

    /// <summary>
    /// One recorded change. Values are kept as written, without formatting or boxing: <see cref="Number"/> holds the
    /// double, the integer, 1 or 0 for a Boolean, or the bits of <see cref="DateTime.ToBinary"/> for a date;
    /// <see cref="Payload"/> holds the text, the decimal text or the <see cref="ExcelStyleOperation"/>.
    /// </summary>
    internal struct ExcelDeferredCellEdit {
        internal int Row;
        internal ushort Column;
        internal ExcelDeferredCellKind Kind;
        internal double Number;
        internal object? Payload;

        internal string Text => (string)Payload!;
        internal ExcelStyleOperation Style => (ExcelStyleOperation)Payload!;
        internal DateTime Date => DateTime.FromBinary(BitConverter.DoubleToInt64Bits(Number));
    }

    /// <summary>
    /// A worksheet's recorded changes in the order they were made, stored in fixed-size chunks so a large write never
    /// copies what it already holds. Writes that arrive row by row, left to right, are already in the order the save
    /// merges them in.
    /// </summary>
    internal sealed class ExcelDeferredCellStore {
        private const int ChunkShift = 12;
        private const int ChunkSize = 1 << ChunkShift;
        private readonly List<ExcelDeferredCellEdit[]> _chunks = new();
        private int _lastRow, _lastColumn;

        internal int Count { get; private set; }

        /// <summary>True while every change came after the previous one in row, then column order.</summary>
        internal bool InCellOrder { get; private set; } = true;

        internal ref ExcelDeferredCellEdit this[int index] => ref _chunks[index >> ChunkShift][index & (ChunkSize - 1)];

        internal void Add(in ExcelDeferredCellEdit edit) {
            if (edit.Row < _lastRow || edit.Row == _lastRow && edit.Column < _lastColumn) InCellOrder = false;
            _lastRow = edit.Row;
            _lastColumn = edit.Column;
            int slot = Count & (ChunkSize - 1);
            if (slot == 0) _chunks.Add(new ExcelDeferredCellEdit[Count == 0 ? 64 : ChunkSize]);
            ExcelDeferredCellEdit[] chunk = _chunks[_chunks.Count - 1];
            if (slot >= chunk.Length) {
                Array.Resize(ref chunk, ChunkSize);
                _chunks[_chunks.Count - 1] = chunk;
            }
            chunk[slot] = edit;
            Count++;
        }

        /// <summary>The changes in row, column and then recording order, as positions in the store; null when they already are.</summary>
        internal int[]? CellOrder() {
            if (InCellOrder) return null;
            // Row (20 bits), column (14 bits) and position (30 bits) pack into one key, so a primitive sort keeps
            // changes to one cell in the order they were made.
            var keys = new ulong[Count];
            for (int index = 0; index < keys.Length; index++) {
                ref ExcelDeferredCellEdit edit = ref this[index];
                keys[index] = ((ulong)(uint)(edit.Row - 1) << 44) | ((ulong)(edit.Column - 1u) << 30) | (uint)index;
            }
            Array.Sort(keys);
            var order = new int[keys.Length];
            for (int index = 0; index < keys.Length; index++) order[index] = (int)(keys[index] & 0x3FFFFFFFUL);
            return order;
        }
    }

    /// <summary>
    /// Plain cell values and styles written to a worksheet whose workbook has not opened its SDK package are recorded
    /// here and written by the owned save. Anything the recorder does not cover opens the SDK, which first applies the
    /// recorded changes through the model in the order they were made.
    /// </summary>
    public partial class ExcelSheet {
        private bool CanDeferCellValue(int row, int column) =>
            _worksheetPartValue == null && _deferredEntry != null && _excelDocument.IsSdkDeferred
            && row >= 1 && row <= ExcelA1.MaxRows && column >= 1 && column <= ExcelA1.MaxColumns
            && _excelDocument.FileOpenAccess != FileAccess.Read;

        private bool TryDeferCellValue(int row, int column, string? value) {
            if (string.IsNullOrEmpty(value) || !CanDeferCellValue(row, column)) return false;
            // A line break needs the wrap format, which the model applies.
            if (value!.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0) return false;
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            return Record(row, column, ExcelDeferredCellKind.Text, 0D, value);
        }

        private bool TryDeferCellValue(int row, int column, double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && CanDeferCellValue(row, column)
            && Record(row, column, ExcelDeferredCellKind.Number, value, null);

        private bool TryDeferCellValue(int row, int column, int value) =>
            CanDeferCellValue(row, column) && Record(row, column, ExcelDeferredCellKind.Integer, value, null);

        private bool TryDeferCellValue(int row, int column, long value) {
            if (!CanDeferCellValue(row, column)) return false;
            // Past 15 significant digits the model keeps the exact value as text.
            return ExcelExactNumber.FitsNumberCell(value)
                ? Record(row, column, ExcelDeferredCellKind.Integer, value, null)
                : Record(row, column, ExcelDeferredCellKind.Text, 0D, InvariantNumberText.Get(value));
        }

        private bool TryDeferCellValue(int row, int column, decimal value) {
            if (!CanDeferCellValue(row, column)) return false;
            string text = value.ToString(CultureInfo.InvariantCulture);
            return Record(row, column, ExcelExactNumber.FitsNumberCell(value) ? ExcelDeferredCellKind.Decimal : ExcelDeferredCellKind.Text, 0D, text);
        }

        private bool TryDeferCellValue(int row, int column, bool value) =>
            CanDeferCellValue(row, column) && Record(row, column, ExcelDeferredCellKind.Boolean, value ? 1D : 0D, null);

        private bool TryDeferCellValue(int row, int column, DateTime value) =>
            CanDeferCellValue(row, column)
            && Record(row, column, ExcelDeferredCellKind.Date, BitConverter.Int64BitsToDouble(value.ToBinary()), null);

        private bool TryDeferBoxedCellValue(int row, int column, object? value) => value switch {
            string text => TryDeferCellValue(row, column, text),
            double number => TryDeferCellValue(row, column, number),
            int number => TryDeferCellValue(row, column, number),
            long number => TryDeferCellValue(row, column, number),
            decimal number => TryDeferCellValue(row, column, number),
            bool flag => TryDeferCellValue(row, column, flag),
            DateTime date => TryDeferCellValue(row, column, date),
            _ => false
        };

        /// <summary>Records a style change for the owned save; false when the model must apply it.</summary>
        private bool TryDeferCellStyle(int row, int column, ExcelStyleOperation operation) =>
            CanDeferCellValue(row, column) && Record(row, column, ExcelDeferredCellKind.Style, 0D, operation);

        private bool TryDeferFontName(int row, int column, string fontName) {
            if (!CanDeferCellValue(row, column)) return false;
            ExcelTextValidation.ThrowIfUnpairedSurrogate(fontName, nameof(fontName), "Font name");
            return TryDeferCellStyle(row, column, ExcelStyleOperation.FontName(fontName));
        }

        private bool TryDeferNumberFormat(int row, int column, string numberFormat) {
            if (numberFormat == null || !CanDeferCellValue(row, column)) return false;
            ExcelTextValidation.ThrowIfUnpairedSurrogate(numberFormat, nameof(numberFormat), "Number format");
            return TryDeferCellStyle(row, column, ExcelStyleOperation.NumberFormat(numberFormat));
        }

        private bool Record(int row, int column, ExcelDeferredCellKind kind, double number, object? payload) {
            var edit = new ExcelDeferredCellEdit { Row = row, Column = (ushort)column, Kind = kind, Number = number, Payload = payload };
            return _excelDocument.RecordDeferredCellEdit(this, edit);
        }

        /// <summary>Applies the recorded changes through the model, as the methods they came from would have.</summary>
        internal void ReplayDeferredCells(ExcelDeferredCellStore store) {
            for (int index = 0; index < store.Count; index++) {
                ref ExcelDeferredCellEdit edit = ref store[index];
                int row = edit.Row, column = edit.Column;
                switch (edit.Kind) {
                    case ExcelDeferredCellKind.Text: CellValue(row, column, edit.Text); break;
                    case ExcelDeferredCellKind.Number: CellValue(row, column, edit.Number); break;
                    case ExcelDeferredCellKind.Integer: CellValue(row, column, (long)edit.Number); break;
                    case ExcelDeferredCellKind.Decimal: CellValue(row, column, decimal.Parse(edit.Text, NumberStyles.Number, CultureInfo.InvariantCulture)); break;
                    case ExcelDeferredCellKind.Boolean: CellValue(row, column, edit.Number != 0D); break;
                    case ExcelDeferredCellKind.Date: CellValue(row, column, edit.Date); break;
                    default: edit.Style.Replay(this, row, column); break;
                }
            }
        }
    }
}
