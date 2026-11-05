using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Keeps one 32-row BIFF block while assembling an owned worksheet record stream.</summary>
    private sealed class StreamingSheet : ExcelTabularRowSink {
        private readonly int _columns;
        private readonly ushort[] _styles;
        private readonly LegacyXlsDirectSharedStringBuilder _strings;
        private readonly DirectTabularWorkbookStream _cells = new(16 * 1024);
        private readonly DirectTabularWorkbookStream _block = new(16 * 1024);
        private readonly int[] _firstCells = new int[DirectRowsPerDbCellBlock];
        private readonly List<uint> _dbCells = new();
        private int _row;
        private int _rows;
        private int _blockRows;
        private bool _complete;

        internal StreamingSheet(int columns, ushort[] styles, LegacyXlsDirectSharedStringBuilder strings) {
            _columns = columns; _styles = styles; _strings = strings;
        }
        internal long Length => _cells.Length;
        internal override void BeginRow(int oneBasedRow) {
            _row = oneBasedRow - 1;
            _firstCells[_blockRows] = checked((int)_block.Length);
        }
        internal override void EndRow() {
            _rows++;
            if (++_blockRows == DirectRowsPerDbCellBlock) FlushBlock();
        }
        internal override void WriteText(string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            uint index = _strings.Add(value);
            int start = PrepareCell();
            int position = start;
            WriteCellHeader(ref position, 0x00fd, 10);
            WriteUInt32(_block.Buffer, position, index);
            position += 4;
            CommitCell(start, position);
        }
        internal override void WriteNumber(double value) {
            ValidateNumber(value);
            int start = PrepareCell();
            int position = start;
            if (BiffRkNumberReader.TryEncodeRkNumber(value, out uint rk)) {
                WriteCellHeader(ref position, 0x027e, 10);
                WriteUInt32(_block.Buffer, position, rk);
                position += 4;
            } else {
                WriteCellHeader(ref position, 0x0203, 14);
                ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
                WriteUInt32(_block.Buffer, position, unchecked((uint)bits));
                WriteUInt32(_block.Buffer, position + 4, unchecked((uint)(bits >> 32)));
                position += 8;
            }
            CommitCell(start, position);
        }
        internal override void WriteBoolean(bool value) {
            int start = PrepareCell();
            int position = start;
            WriteCellHeader(ref position, 0x0205, 8);
            _block.Buffer[position++] = value ? (byte)1 : (byte)0;
            _block.Buffer[position++] = 0;
            CommitCell(start, position);
        }
        internal override void WriteBlank() {
            int start = PrepareCell();
            int position = start;
            WriteCellHeader(ref position, 0x0201, 6);
            CommitCell(start, position);
        }
        private int PrepareCell() {
            int start = checked((int)_block.Position);
            _block.EnsureCapacity(checked(start + 18));
            _block.RegisterDirectWriteUpperBound(start + 18);
            return start;
        }
        private void WriteCellHeader(ref int position, ushort type, ushort length) =>
            WriteDirectFixedCellHeader(_block.Buffer, ref position, type, length,
                checked((ushort)_row), checked((ushort)Column), _styles[checked((int)Style)]);
        private void CommitCell(int start, int end) => _block.Write(_block.Buffer, start, end - start);

        private void FlushBlock() {
            if (_blockRows == 0) return;
            int start = checked((int)_cells.Position);
            int end = checked(start + _blockRows * 20 + (int)_block.Length + 8 + _blockRows * 2);
            _cells.EnsureCapacity(end);
            _cells.RegisterDirectWriteUpperBound(end);
            byte[] output = _cells.Buffer;
            int position = start;
            int firstRow = _rows - _blockRows;
            for (int row = 0; row < _blockRows; row++)
                WriteDirectRowRecord(output, ref position, checked((ushort)(firstRow + row)), 0, checked((ushort)(_columns - 1)));
            int cellsStart = position;
            Buffer.BlockCopy(_block.Buffer, 0, output, position, checked((int)_block.Length));
            position += checked((int)_block.Length);
            for (int row = 0; row < _blockRows; row++) _firstCells[row] += cellsStart;
            int dbCell = position;
            _dbCells.Add(checked((uint)dbCell));
            WriteDirectDbCellRecord(output, ref position, start, _firstCells.AsSpan(0, _blockRows), dbCell);
            _cells.Write(output, start, position - start);
            _block.SetLength(0);
            _block.Position = 0;
            _blockRows = 0;
        }
        internal override void Complete() { FlushBlock(); _complete = true; }
        internal void WriteTo(DirectTabularWorkbookStream workbook) {
            if (!_complete) throw new InvalidOperationException("Complete each XLS table before publishing the workbook.");
            var dimensions = _rows == 0 ? default : new LegacyXlsDimensions(0, checked((uint)_rows), 0, checked((ushort)_columns));
            byte[] index = new byte[checked(16 + _dbCells.Count * 4)];
            WriteUInt32(index, 8, checked((uint)_rows));
            WriteDirectWorksheetPrefix(workbook, dimensions, index, out long indexPosition, out long widthPosition);
            uint offset = checked((uint)workbook.Position);
            workbook.Write(_cells.Buffer, 0, checked((int)_cells.Length));
            WriteDirectWorksheetSuffix(workbook);
            PatchIndexRecord(workbook, indexPosition, widthPosition, _dbCells.Select(position => checked(position + offset)).ToArray());
        }
        public override void Dispose() { _block.Dispose(); _cells.Dispose(); }
    }
}
