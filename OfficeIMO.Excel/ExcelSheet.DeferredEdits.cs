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
        /// <summary>
        /// A calculated formula result: the cell keeps its formula and style and takes the value, the type and the
        /// calculate-on-load flag as they are written.
        /// </summary>
        FormulaResult,
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

        /// <summary>A formula result: its value text (null for none), its type and its calculate-on-load flag.</summary>
        internal static ExcelDeferredCellEdit FormulaResult(int row, int column, string? text, ExcelCellValueType? type, bool? calculate) {
            int code = (type.HasValue ? (int)type.Value + 1 : 0) | (calculate.HasValue ? (calculate.Value ? 1 : 2) << 4 : 0);
            return new ExcelDeferredCellEdit {
                Row = row, Column = checked((ushort)column), Kind = ExcelDeferredCellKind.FormulaResult, Number = code, Payload = text
            };
        }

        internal string? FormulaResultText => (string?)Payload;

        internal ExcelCellValueType? FormulaResultType => ((int)Number & 0x0F) is int type && type != 0 ? (ExcelCellValueType)(type - 1) : null;

        internal bool? FormulaResultCalculate => (((int)Number >> 4) & 0x03) switch { 1 => true, 2 => false, _ => null };
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
        // Plain values written row by row to a new, empty sheet are held as the cell buffer the tabular writer saves;
        // anything else turns them into recorded changes first.
        private ExcelCellValueBuffer? _tabular;

        internal ExcelDeferredCellStore(bool tabular = false) {
            if (tabular) _tabular = new ExcelCellValueBuffer();
        }

        internal int Count { get; private set; }

        /// <summary>The parts of the tables whose header text these changes rename, or null when none do.</summary>
        internal HashSet<string>? HeaderTables { get; set; }

        /// <summary>True while every change came after the previous one in row, then column order.</summary>
        internal bool InCellOrder { get; private set; } = true;

        /// <summary>The cells, while every change has been a plain value written after the previous one; null otherwise.</summary>
        internal ExcelCellValueBuffer? Tabular => _tabular;

        /// <summary>Whether <paramref name="edit"/> keeps the store's cells in the tabular buffer.</summary>
        internal bool TakesTabular(in ExcelDeferredCellEdit edit) =>
            _tabular != null && edit.Kind != ExcelDeferredCellKind.Style && _tabular.CanAdd(edit.Row, edit.Column);

        internal ref ExcelDeferredCellEdit this[int index] {
            get {
                if (_tabular != null) LeaveTabular();
                return ref _chunks[index >> ChunkShift][index & (ChunkSize - 1)];
            }
        }

        // Runs for every recorded cell, so it is compiled optimized from the first call rather than tiered up.
#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        internal void Add(in ExcelDeferredCellEdit edit) {
            if (_tabular != null) {
                if (_tabular.CanAdd(edit.Row, edit.Column) && TryBuffer(edit, out ExcelBufferedCell cell)) {
                    _tabular.Add(edit.Row, edit.Column, cell);
                    _lastRow = edit.Row;
                    _lastColumn = edit.Column;
                    Count++;
                    return;
                }
                LeaveTabular();
            }
            AddEdit(edit);
        }

        // Runs for every recorded cell, so it is compiled optimized from the first call rather than tiered up.
#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private static bool TryBuffer(in ExcelDeferredCellEdit edit, out ExcelBufferedCell cell) {
            switch (edit.Kind) {
                case ExcelDeferredCellKind.Text: cell = ExcelBufferedCell.Text(edit.Text); return true;
                case ExcelDeferredCellKind.Number: cell = ExcelBufferedCell.Number(edit.Number); return true;
                case ExcelDeferredCellKind.Integer:
                    cell = edit.Number >= int.MinValue && edit.Number <= int.MaxValue ? ExcelBufferedCell.Int32((int)edit.Number) : ExcelBufferedCell.Int64((long)edit.Number);
                    return true;
                case ExcelDeferredCellKind.Decimal:
                    cell = ExcelBufferedCell.FromValue(decimal.Parse(edit.Text, NumberStyles.Number, CultureInfo.InvariantCulture));
                    return true;
                case ExcelDeferredCellKind.Boolean: cell = ExcelBufferedCell.Boolean(edit.Number != 0D); return true;
                case ExcelDeferredCellKind.Date: cell = ExcelBufferedCell.DateTimeCell(edit.Date); return true;
                default: cell = default; return false;
            }
        }

        /// <summary>Turns the buffered cells into recorded changes, in the order they were written.</summary>
        private void LeaveTabular() {
            ExcelCellValueBuffer buffer = _tabular!;
            _tabular = null;
            Count = 0;
            _lastRow = _lastColumn = 0;
            for (int written = 0; written < buffer.WrittenRowCount; written++) {
                int row = buffer.GetRowNumber(written);
                for (int index = buffer.GetFirstCellIndex(written); index < buffer.GetEndCellIndex(written); index++) {
                    ExcelBufferedCell cell = buffer.GetCell(index);
                    var edit = new ExcelDeferredCellEdit { Row = row, Column = (ushort)buffer.GetColumn(index) };
                    switch (cell.Kind) {
                        case ExcelBufferedCellKind.Text: edit.Kind = ExcelDeferredCellKind.Text; edit.Payload = cell.Reference; break;
                        case ExcelBufferedCellKind.Number: edit.Kind = ExcelDeferredCellKind.Number; edit.Number = cell.NumberValue; break;
                        case ExcelBufferedCellKind.Int32:
                        case ExcelBufferedCellKind.Int64: edit.Kind = ExcelDeferredCellKind.Integer; edit.Number = cell.Bits; break;
                        case ExcelBufferedCellKind.Boolean: edit.Kind = ExcelDeferredCellKind.Boolean; edit.Number = cell.Bits != 0 ? 1D : 0D; break;
                        case ExcelBufferedCellKind.DateTime:
                            edit.Kind = ExcelDeferredCellKind.Date;
                            edit.Number = BitConverter.Int64BitsToDouble(cell.DateTimeValue.ToBinary());
                            break;
                        default:
                            edit.Kind = ExcelDeferredCellKind.Decimal;
                            edit.Payload = ((decimal)cell.Reference!).ToString(CultureInfo.InvariantCulture);
                            break;
                    }
                    AddEdit(edit);
                }
            }
        }

        private void AddEdit(in ExcelDeferredCellEdit edit) {
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
            if (_tabular != null) LeaveTabular();
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
        // The recording path runs once per cell, so it is compiled optimized from the first call: a process that
        // writes a few workbooks would otherwise spend most of them in the unoptimized tier.
#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool CanDeferCellValue(int row, int column) =>
            _worksheetPartValue == null && _deferredEntry != null && _excelDocument.IsSdkDeferred
            && row >= 1 && row <= ExcelA1.MaxRows && column >= 1 && column <= ExcelA1.MaxColumns
            && _excelDocument.FileOpenAccess != FileAccess.Read;

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, string? value) {
            if (string.IsNullOrEmpty(value) || !CanDeferCellValue(row, column)) return false;
            // A line break needs the wrap format, which the model applies.
            if (value!.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0) return false;
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            return Record(row, column, ExcelDeferredCellKind.Text, 0D, value);
        }

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && CanDeferCellValue(row, column)
            && Record(row, column, ExcelDeferredCellKind.Number, value, null);

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, int value) =>
            CanDeferCellValue(row, column) && Record(row, column, ExcelDeferredCellKind.Integer, value, null);

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, long value) {
            if (!CanDeferCellValue(row, column)) return false;
            // Past 15 significant digits the model keeps the exact value as text.
            return ExcelExactNumber.FitsNumberCell(value)
                ? Record(row, column, ExcelDeferredCellKind.Integer, value, null)
                : Record(row, column, ExcelDeferredCellKind.Text, 0D, InvariantNumberText.Get(value));
        }

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, decimal value) {
            if (!CanDeferCellValue(row, column)) return false;
            string text = value.ToString(CultureInfo.InvariantCulture);
            return Record(row, column, ExcelExactNumber.FitsNumberCell(value) ? ExcelDeferredCellKind.Decimal : ExcelDeferredCellKind.Text, 0D, text);
        }

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
        private bool TryDeferCellValue(int row, int column, bool value) =>
            CanDeferCellValue(row, column) && Record(row, column, ExcelDeferredCellKind.Boolean, value ? 1D : 0D, null);

#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
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

        // Runs for every recorded cell, so it is compiled optimized from the first call rather than tiered up.
#if NETCOREAPP3_0_OR_GREATER
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
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
                    case ExcelDeferredCellKind.FormulaResult: ReplayFormulaResult(edit); break;
                    default: edit.Style.Replay(this, row, column); break;
                }
            }
        }

        private void ReplayFormulaResult(in ExcelDeferredCellEdit edit) {
            ExcelFormulaCellRecord cell = FormulaModel.GetOrCreate(edit.Row, edit.Column);
            cell.ClearInlineString();
            cell.SetValue(edit.FormulaResultText, edit.FormulaResultType);
            if (cell.HasFormula) cell.SetCalculateCell(edit.FormulaResultCalculate);
        }
    }
}
