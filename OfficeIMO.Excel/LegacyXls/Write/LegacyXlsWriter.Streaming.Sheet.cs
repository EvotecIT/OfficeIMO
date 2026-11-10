using System.Buffers;
using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Keeps one 32-row BIFF block while appending completed blocks to owned, pooled chunks.</summary>
    private sealed class StreamingSheet : ExcelTabularRowSink {
        private const int ChunkSize = 1 << 20;
        // BIFF8 LABEL stores at most 255 characters; longer text uses the shared-string table.
        private const int MaximumLabelLength = 255;
        private readonly int _columns;
        private readonly ushort[] _styles;
        private readonly LegacyXlsDirectSharedStringBuilder _strings;
        private readonly bool _inlineText;
        private readonly DirectTabularWorkbookStream _block = new(16 * 1024);
        private readonly int[] _firstCells = new int[DirectRowsPerDbCellBlock];
        private readonly List<uint> _dbCells = new();
        private readonly List<ArraySegment<byte>> _chunks = new();
        private byte[]? _chunk;
        private int _chunkUsed;
        private long _length;
        private int _row;
        private int _rows;
        private int _blockRows;
        private int _rkRunStart = -1;
        private int _rkRunCount;
        private int _rkLastColumn;
        private bool _complete;

        internal StreamingSheet(int columns, ushort[] styles, LegacyXlsDirectSharedStringBuilder strings, bool inlineText) {
            _columns = columns; _styles = styles; _strings = strings; _inlineText = inlineText;
        }
        internal override void BeginRow(int oneBasedRow) {
            _rkRunStart = -1;
            _row = oneBasedRow - 1;
            _firstCells[_blockRows] = checked((int)_block.Length);
        }
        internal override void EndRow() {
            _rows++;
            if (++_blockRows == DirectRowsPerDbCellBlock) FlushBlock();
        }
        internal override void WriteText(string value) {
            _rkRunStart = -1;
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            if (_inlineText && value.Length <= MaximumLabelLength) {
                WriteLabel(value);
                return;
            }
            uint index = _strings.Add(value);
            int position = PrepareCell(18);
            WriteCellHeader(ref position, 0x00fd, 10);
            WriteUInt32(_block.Buffer, position, index);
            position += 4;
            _block.CommitDirectWrite(position);
        }
        /// <summary>Writes a BIFF8 LABEL record with an 8-bit string when every character fits Latin-1.</summary>
        private void WriteLabel(string value) {
            bool compressed = CanUseCompressedString(value);
            int textBytes = compressed ? value.Length : value.Length * 2;
            int position = PrepareCell(13 + textBytes);
            WriteCellHeader(ref position, 0x0204, checked((ushort)(9 + textBytes)));
            byte[] buffer = _block.Buffer;
            WriteUInt16(buffer, position, checked((ushort)value.Length));
            buffer[position + 2] = compressed ? (byte)0 : (byte)1;
            position += 3;
            if (compressed) {
                for (int index = 0; index < value.Length; index++) buffer[position + index] = unchecked((byte)value[index]);
            } else {
                System.Text.Encoding.Unicode.GetBytes(value, 0, value.Length, buffer, position);
            }
            _block.CommitDirectWrite(position + textBytes);
        }
        internal override void WriteNumber(double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                // BIFF8 has no NaN or infinite number; Excel shows #NUM! (error 0x24) for such a result.
                _rkRunStart = -1;
                int errorPosition = PrepareCell(18);
                WriteCellHeader(ref errorPosition, 0x0205, 8);
                _block.Buffer[errorPosition++] = 0x24;
                _block.Buffer[errorPosition++] = 1;
                _block.CommitDirectWrite(errorPosition);
                return;
            }
            int position = PrepareCell(18);
            if (BiffRkNumberReader.TryEncodeRkNumber(value, out uint rk)) {
                WriteRkCell(ref position, rk);
            } else {
                _rkRunStart = -1;
                WriteCellHeader(ref position, 0x0203, 14);
                ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
                WriteUInt32(_block.Buffer, position, unchecked((uint)bits));
                WriteUInt32(_block.Buffer, position + 4, unchecked((uint)(bits >> 32)));
                position += 8;
            }
            _block.CommitDirectWrite(position);
        }
        internal override void WriteBoolean(bool value) {
            _rkRunStart = -1;
            int position = PrepareCell(18);
            WriteCellHeader(ref position, 0x0205, 8);
            _block.Buffer[position++] = value ? (byte)1 : (byte)0;
            _block.Buffer[position++] = 0;
            _block.CommitDirectWrite(position);
        }
        internal override void WriteBlank() {
            _rkRunStart = -1;
            int position = PrepareCell(18);
            WriteCellHeader(ref position, 0x0201, 6);
            _block.CommitDirectWrite(position);
        }
        /// <summary>Extends adjacent RK cells in the owned row block without deferring any values.</summary>
        private void WriteRkCell(ref int position, uint rk) {
            if (_rkRunStart >= 0 && Column == _rkLastColumn + 1) {
                _ = checked((ushort)_row);
                ushort column = checked((ushort)Column);
                ushort style = _styles[checked((int)Style)];
                // A multi-cell record already ends with colLast; replace that trailer.
                if (_rkRunCount > 1) position -= 2;
                WriteUInt16(_block.Buffer, position, style);
                WriteUInt32(_block.Buffer, position + 2, rk);
                WriteUInt16(_block.Buffer, position + 6, column);
                position += 8;
                WriteUInt16(_block.Buffer, _rkRunStart, 0x00bd);
                WriteUInt16(_block.Buffer, _rkRunStart + 2, checked((ushort)(position - _rkRunStart - 4)));
                _rkRunCount++;
            } else {
                _rkRunStart = position;
                _rkRunCount = 1;
                WriteCellHeader(ref position, 0x027e, 10);
                WriteUInt32(_block.Buffer, position, rk);
                position += 4;
            }
            _rkLastColumn = Column;
        }

        private int PrepareCell(int maximumBytes) {
            int start = checked((int)_block.Position);
            _block.EnsureCapacity(checked(start + maximumBytes));
            _block.RegisterDirectWriteUpperBound(start + maximumBytes);
            return start;
        }
        private void WriteCellHeader(ref int position, ushort type, ushort length) =>
            WriteDirectFixedCellHeader(_block.Buffer, ref position, type, length,
                checked((ushort)_row), checked((ushort)Column), _styles[checked((int)Style)]);

        /// <summary>Appends the block's ROW records, cells and DBCELL; a block never spans two chunks.</summary>
        private void FlushBlock() {
            if (_blockRows == 0) return;
            int size = checked(_blockRows * 20 + (int)_block.Length + 8 + _blockRows * 2);
            if (_chunk == null || _chunk.Length - _chunkUsed < size) StartChunk(size);
            byte[] output = _chunk!;
            int start = _chunkUsed;
            int position = start;
            int firstRow = _rows - _blockRows;
            for (int row = 0; row < _blockRows; row++)
                WriteDirectRowRecord(output, ref position, checked((ushort)(firstRow + row)), 0, checked((ushort)(_columns - 1)));
            int cellsStart = position;
            Buffer.BlockCopy(_block.Buffer, 0, output, position, checked((int)_block.Length));
            position += checked((int)_block.Length);
            for (int row = 0; row < _blockRows; row++) _firstCells[row] += cellsStart;
            int dbCell = position;
            // DBCELL offsets are relative to the sheet's cell table; chunk gaps are not part of it.
            _dbCells.Add(checked((uint)(_length + (dbCell - start))));
            WriteDirectDbCellRecord(output, ref position, start, _firstCells.AsSpan(0, _blockRows), dbCell);
            _length += position - start;
            _chunkUsed = position;
            _block.SetLength(0);
            _block.Position = 0;
            _blockRows = 0;
        }
        private void StartChunk(int minimumSize) {
            if (_chunk != null) _chunks.Add(new ArraySegment<byte>(_chunk, 0, _chunkUsed));
            _chunk = ArrayPool<byte>.Shared.Rent(Math.Max(ChunkSize, minimumSize));
            _chunkUsed = 0;
        }
        internal override void Complete() {
            FlushBlock();
            if (_chunk != null) {
                _chunks.Add(new ArraySegment<byte>(_chunk, 0, _chunkUsed));
                _chunk = null;
            }
            _complete = true;
        }
        /// <summary>Adds this sheet's records at <paramref name="position"/> in the workbook stream and returns the following position.</summary>
        internal long AddSegments(List<ArraySegment<byte>> segments, long position) {
            if (!_complete) throw new InvalidOperationException("Complete each XLS table before publishing the workbook.");
            var dimensions = _rows == 0 ? default : new LegacyXlsDimensions(0, checked((uint)_rows), 0, checked((ushort)_columns));
            byte[] index = new byte[checked(16 + _dbCells.Count * 4)];
            WriteUInt32(index, 8, checked((uint)_rows));
            var prefix = new MemoryStream(checked(512 + index.Length));
            WriteDirectWorksheetPrefix(prefix, dimensions, index, out long indexPosition, out long widthPosition);
            long cellsStart = checked(position + prefix.Length);
            var dbCells = new uint[_dbCells.Count];
            for (int block = 0; block < dbCells.Length; block++) dbCells[block] = checked((uint)(cellsStart + _dbCells[block]));
            PatchIndexRecord(prefix, indexPosition, checked(position + widthPosition), dbCells);
            segments.Add(new ArraySegment<byte>(prefix.GetBuffer(), 0, checked((int)prefix.Length)));
            segments.AddRange(_chunks);
            var suffix = new MemoryStream(64);
            WriteDirectWorksheetSuffix(suffix);
            segments.Add(new ArraySegment<byte>(suffix.GetBuffer(), 0, checked((int)suffix.Length)));
            return checked(cellsStart + _length + suffix.Length);
        }
        public override void Dispose() {
            _block.Dispose();
            if (_chunk != null) _chunks.Add(new ArraySegment<byte>(_chunk, 0, _chunkUsed));
            _chunk = null;
            foreach (ArraySegment<byte> chunk in _chunks) {
                Array.Clear(chunk.Array!, 0, chunk.Count);
                ArrayPool<byte>.Shared.Return(chunk.Array!);
            }
            _chunks.Clear();
        }
    }
}
