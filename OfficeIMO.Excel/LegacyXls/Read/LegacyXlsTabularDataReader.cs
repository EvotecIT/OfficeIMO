using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Projection;
using System.Buffers;
using System.Buffers.Binary;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    /// <summary>
    /// Forward-only BIFF8 worksheet reader that retains one primitive row and does not
    /// materialize legacy cells or create an Open XML workbook projection.
    /// </summary>
    internal sealed partial class LegacyXlsTabularDataReader : DbDataReader {
        private readonly LegacyBiffSource _bytes;
        private readonly byte[]? _bufferedBytes;
        private readonly IReadOnlyList<string> _sharedStrings;
        private readonly bool[] _dateStyles;
        private readonly bool _uses1904DateSystem;
        private readonly ExcelReadOptions _options;
        private readonly CancellationToken _cancellationToken;
        private CancellationToken _activeReadCancellationToken;
        private readonly string[] _headers;
        private readonly Dictionary<string, int> _ordinals;
        private readonly ValueKind[] _kinds;
        private readonly double[] _numbers;
        private readonly bool[] _booleans;
        private readonly string?[] _strings;
        private readonly Type[] _columnTypes;
        private readonly List<LegacyXlsBufferedRow>? _schemaRows;
        private readonly int _firstColumn;
        private readonly int _firstDataRow;
        private readonly int _lastDataRow;
        private int[]? _bufferedRowOffsets;
        private readonly int _bufferedWorksheetEndOffset;
        private int _position;
        private int _nextRow;
        private int _schemaRowIndex;
        private int _recordsSinceCancellationCheck;
        private bool _hasCurrentRow;
        private bool _closed;

        internal LegacyXlsTabularDataReader(
            LegacyBiffSource bytes,
            int sheetOffset,
            IReadOnlyList<string> sharedStrings,
            bool[] dateStyles,
            bool uses1904DateSystem,
            bool hasHeaderRow,
            ExcelReadOptions options,
            CancellationToken cancellationToken) {
            _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
            _bufferedBytes = bytes.ContiguousBuffer;
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _dateStyles = dateStyles ?? throw new ArgumentNullException(nameof(dateStyles));
            _uses1904DateSystem = uses1904DateSystem;
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _cancellationToken = cancellationToken;
            _activeReadCancellationToken = cancellationToken;

            int[]? bufferedRowOffsets = _bufferedBytes != null
                ? ArrayPool<int>.Shared.Rent(ushort.MaxValue + 1)
                : null;
            int bufferedWorksheetEndOffset;
            try {
                Discover(
                    sheetOffset,
                    bufferedRowOffsets,
                    out bufferedWorksheetEndOffset,
                    out int firstRow,
                    out int lastRow,
                    out int firstColumn,
                    out int lastColumn,
                    out Dictionary<int, string?>? headerValues);
                _bufferedRowOffsets = bufferedRowOffsets;
                _bufferedWorksheetEndOffset = bufferedWorksheetEndOffset;
                _firstColumn = firstColumn;
                _firstDataRow = hasHeaderRow && firstRow >= 0 ? checked(firstRow + 1) : firstRow;
                _lastDataRow = lastRow;
                _position = sheetOffset;
                _nextRow = _firstDataRow;

                int fieldCount = lastColumn >= firstColumn
                    ? checked(lastColumn - firstColumn + 1)
                    : 0;
                if (fieldCount > options.MaxDataReaderColumns) {
                    throw new InvalidDataException(
                        $"XLS table column count {fieldCount} exceeds the configured limit of {options.MaxDataReaderColumns}.");
                }
                _headers = ExcelHeaderNameHelper.BuildUniqueHeaders(
                    fieldCount,
                    ordinal => hasHeaderRow
                        && headerValues != null
                        && headerValues.TryGetValue(ordinal + firstColumn, out string? value)
                            ? value
                            : null,
                    options.NormalizeHeaders);
                _ordinals = new Dictionary<string, int>(_headers.Length, StringComparer.OrdinalIgnoreCase);
                for (int index = 0; index < _headers.Length; index++) {
                    _ordinals[_headers[index]] = index;
                }

                _kinds = new ValueKind[fieldCount];
                _numbers = new double[fieldCount];
                _booleans = new bool[fieldCount];
                _strings = new string?[fieldCount];
                _columnTypes = CreateObjectColumnTypes(fieldCount);
                _schemaRows = _options.InferSchema
                    ? BufferSchemaRows()
                    : null;
            } catch {
                if (bufferedRowOffsets != null) {
                    ArrayPool<int>.Shared.Return(bufferedRowOffsets);
                }
                throw;
            }
        }

        public override object this[int ordinal] => GetValue(ordinal);
        public override object this[string name] => GetValue(GetOrdinal(name));
        public override int Depth => 0;
        public override int FieldCount => _headers.Length;
        public override bool HasRows => _firstDataRow >= 0 && _firstDataRow <= _lastDataRow;
        public override bool IsClosed => _closed;
        public override int RecordsAffected => -1;

        public override bool Read() => ReadCore(_cancellationToken);

        public override Task<bool> ReadAsync(CancellationToken cancellationToken) {
            try {
                return Task.FromResult(ReadCore(cancellationToken));
            } catch (OperationCanceledException exception) when (exception.CancellationToken.IsCancellationRequested) {
                return Task.FromCanceled<bool>(exception.CancellationToken);
            } catch (Exception exception) {
                return Task.FromException<bool>(exception);
            }
        }

        private bool ReadCore(CancellationToken cancellationToken) {
            ThrowIfClosed();
            _cancellationToken.ThrowIfCancellationRequested();
            cancellationToken.ThrowIfCancellationRequested();
            _activeReadCancellationToken = cancellationToken;
            try {
                _hasCurrentRow = false;
                if (_schemaRows != null && _schemaRowIndex < _schemaRows.Count) {
                    LoadBufferedRow(_schemaRows[_schemaRowIndex++]);
                    _hasCurrentRow = true;
                    return true;
                }

                bool result = ReadSourceRow();
                ThrowIfReadCancellationRequested();
                return result;
            } finally {
                _activeReadCancellationToken = _cancellationToken;
            }
        }

        private bool ReadSourceRow() {
            if (_nextRow < 0 || _nextRow > _lastDataRow) return false;

            Array.Clear(_kinds, 0, _kinds.Length);
            Array.Clear(_strings, 0, _strings.Length);
            int currentRow = _nextRow++;
            int pendingFormulaOrdinal = -1;
            ushort pendingFormulaStyle = 0;

            if (_bufferedBytes != null) {
                int[] rowOffsets = _bufferedRowOffsets!;
                int rowStartMarker = rowOffsets[currentRow];
                if (rowStartMarker == 0) {
                    _hasCurrentRow = true;
                    return true;
                }
                int rowEnd = _bufferedWorksheetEndOffset;
                for (int nextRow = currentRow + 1; nextRow <= _lastDataRow; nextRow++) {
                    int nextMarker = rowOffsets[nextRow];
                    if (nextMarker != 0) {
                        rowEnd = nextMarker - 1;
                        break;
                    }
                }
                ReadBufferedCurrentRow(
                    _bufferedBytes,
                    rowStartMarker - 1,
                    rowEnd,
                    ref pendingFormulaOrdinal,
                    ref pendingFormulaStyle);
            } else {
                ReadPagedCurrentRow(currentRow, ref pendingFormulaOrdinal, ref pendingFormulaStyle);
            }

            _hasCurrentRow = true;
            return true;
        }

        private void ReadPagedCurrentRow(
            int currentRow,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            while (true) {
                int recordOffset = _position;
                if (!TryReadRecord(_bytes, ref _position, out RecordSlice record)) {
                    if (pendingFormulaOrdinal >= 0) {
                        throw MissingFormulaString();
                    }
                    break;
                }
                CheckCancellation();

                if (pendingFormulaOrdinal >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) {
                        throw MissingFormulaString();
                    }
                    StoreFormulaString(record, pendingFormulaOrdinal, pendingFormulaStyle, ref _position);
                    pendingFormulaOrdinal = -1;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Eof) break;

                if (!TryGetCellBounds(record, out int row, out int firstColumn, out _)) continue;
                if (row < currentRow) continue;
                if (row > currentRow) {
                    _position = recordOffset;
                    break;
                }

                StoreCellRecord(record, ref pendingFormulaOrdinal, ref pendingFormulaStyle);
            }
        }

        private void ReadBufferedCurrentRow(
            byte[] bytes,
            int rowStart,
            int rowEnd,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            bool checkCancellation = CanCancelCurrentRead;
            _position = rowStart;
            while (_position < rowEnd) {
                int recordOffset = _position;
                ushort type = (ushort)(bytes[recordOffset] | bytes[recordOffset + 1] << 8);
                int length = bytes[recordOffset + 2] | bytes[recordOffset + 3] << 8;
                int payloadOffset = recordOffset + 4;
                _position = payloadOffset + length;
                if (checkCancellation) {
                    CheckCancellation();
                }

                if (pendingFormulaOrdinal >= 0) {
                    if (type != (ushort)BiffRecordType.String) {
                        throw MissingFormulaString();
                    }
                    StoreFormulaString(
                        new RecordSlice(type, recordOffset, payloadOffset, length),
                        pendingFormulaOrdinal,
                        pendingFormulaStyle,
                        ref _position);
                    pendingFormulaOrdinal = -1;
                    continue;
                }
                if (type == (ushort)BiffRecordType.Eof) break;

                if (!IsCellRecordType(type)) continue;
                StoreBufferedCellRecord(
                    bytes,
                    type,
                    recordOffset,
                    payloadOffset,
                    length,
                    ref pendingFormulaOrdinal,
                    ref pendingFormulaStyle);
            }
            if (pendingFormulaOrdinal >= 0) {
                throw MissingFormulaString();
            }
        }

        public override bool NextResult() => false;

        public override void Close() {
            _closed = true;
            _hasCurrentRow = false;
            int[]? rowOffsets = Interlocked.Exchange(ref _bufferedRowOffsets, null);
            if (rowOffsets != null) {
                ArrayPool<int>.Shared.Return(rowOffsets);
            }
        }

        protected override void Dispose(bool disposing) {
            if (disposing) Close();
            base.Dispose(disposing);
        }

        private void Discover(
            int sheetOffset,
            int[]? bufferedRowOffsets,
            out int bufferedWorksheetEndOffset,
            out int firstRow,
            out int lastRow,
            out int firstColumn,
            out int lastColumn,
            out Dictionary<int, string?>? headerValues) {
            int offset = sheetOffset;
            firstRow = int.MaxValue;
            lastRow = -1;
            firstColumn = int.MaxValue;
            lastColumn = -1;
            bool sawBof = false;
            bool sawEof = false;
            int pendingHeaderColumn = -1;
            int previousCellRow = -1;
            bool checkCancellation = CanCancelCurrentRead;
            bufferedWorksheetEndOffset = -1;
            headerValues = null;

            while (TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (checkCancellation) {
                    CheckCancellation();
                }
                if (!sawBof) {
                    if (record.Type != (ushort)BiffRecordType.Bof || record.Length < 4) {
                        throw new InvalidDataException("The XLS worksheet stream is missing a valid BOF record.");
                    }
                    ushort version = ReadUInt16(_bytes, record.PayloadOffset);
                    ushort substreamType = ReadUInt16(_bytes, record.PayloadOffset + 2);
                    if (version != 0x0600 || substreamType != 0x0010) {
                        throw new InvalidDataException("The selected XLS substream is not a BIFF8 worksheet.");
                    }
                    sawBof = true;
                    continue;
                }
                if (pendingHeaderColumn >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) {
                        throw MissingFormulaString();
                    }
                    headerValues![pendingHeaderColumn] = ReadFormulaStringValue(record, ref offset);
                    pendingHeaderColumn = -1;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Eof) {
                    sawEof = true;
                    bufferedWorksheetEndOffset = record.Offset;
                    break;
                }
                if (!TryGetCellBounds(record, out int row, out int recordFirstColumn, out int recordLastColumn)) {
                    continue;
                }
                if (row < previousCellRow) {
                    throw new InvalidDataException(
                        $"The XLS worksheet contains decreasing cell row index {row} after row {previousCellRow}.");
                }
                if (bufferedRowOffsets != null && row != previousCellRow) {
                    int firstUninitializedRow = previousCellRow < 0 ? row : previousCellRow + 1;
                    for (int missingRow = firstUninitializedRow; missingRow < row; missingRow++) {
                        bufferedRowOffsets[missingRow] = 0;
                    }
                    bufferedRowOffsets[row] = record.Offset + 1;
                }
                previousCellRow = row;

                if (row < firstRow) {
                    firstRow = row;
                    headerValues = new Dictionary<int, string?>();
                }
                if (row == firstRow) {
                    ReadHeaderCells(record, headerValues!);
                    if (record.Type == (ushort)BiffRecordType.Formula && FormulaExpectsString(record)) {
                        pendingHeaderColumn = recordFirstColumn;
                    }
                }
                lastRow = row;
                firstColumn = Math.Min(firstColumn, recordFirstColumn);
                lastColumn = Math.Max(lastColumn, recordLastColumn);
            }

            if (!sawBof || !sawEof) {
                throw new InvalidDataException("The XLS worksheet substream is truncated before EOF.");
            }
            if (firstRow == int.MaxValue) {
                firstRow = -1;
                firstColumn = 0;
            }
        }

        private bool TryGetCellBounds(
            RecordSlice record,
            out int row,
            out int firstColumn,
            out int lastColumn) {
            row = -1;
            firstColumn = -1;
            lastColumn = -1;
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.Blank:
                case BiffRecordType.BoolErr:
                case BiffRecordType.Formula:
                case BiffRecordType.Label:
                case BiffRecordType.LabelSst:
                case BiffRecordType.Number:
                case BiffRecordType.Rk:
                    if (record.Length < 6) throw TruncatedCell(record);
                    row = ReadDiscoveryUInt16(record.PayloadOffset);
                    firstColumn = ReadDiscoveryUInt16(record.PayloadOffset + 2);
                    lastColumn = firstColumn;
                    return true;
                case BiffRecordType.MulBlank:
                case BiffRecordType.MulRk:
                    if (record.Length < 6) throw TruncatedCell(record);
                    row = ReadDiscoveryUInt16(record.PayloadOffset);
                    firstColumn = ReadDiscoveryUInt16(record.PayloadOffset + 2);
                    lastColumn = ReadDiscoveryUInt16(record.PayloadOffset + record.Length - 2);
                    if (lastColumn < firstColumn) throw new InvalidDataException("A BIFF multiple-cell record has an invalid column range.");
                    int cellCount = checked(lastColumn - firstColumn + 1);
                    int bytesPerCell = record.Type == (ushort)BiffRecordType.MulBlank ? 2 : 6;
                    int expectedLength = checked(4 + cellCount * bytesPerCell + 2);
                    if (record.Length != expectedLength) {
                        string recordName = record.Type == (ushort)BiffRecordType.MulBlank ? "MULBLANK" : "MULRK";
                        throw new InvalidDataException($"A {recordName} record has an invalid payload length.");
                    }
                    return true;
                default:
                    return false;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool TryReadDiscoveryRecord(ref int offset, out RecordSlice record) {
            byte[]? bytes = _bufferedBytes;
            if (bytes == null) {
                return TryReadRecord(_bytes, ref offset, out record);
            }

            int sourceLength = _bytes.Length;
            if (offset == sourceLength) {
                record = default;
                return false;
            }
            if (offset < 0 || offset > sourceLength - 4) {
                throw new InvalidDataException("The BIFF stream ended inside a record header.");
            }

            int recordOffset = offset;
            ushort type = (ushort)(bytes[offset] | bytes[offset + 1] << 8);
            int length = bytes[offset + 2] | bytes[offset + 3] << 8;
            int payloadOffset = offset + 4;
            if (length > sourceLength - payloadOffset) {
                throw new InvalidDataException(
                    $"BIFF record 0x{type:X4} at offset {recordOffset} declares {length} payload bytes, but the stream ends early.");
            }

            offset = payloadOffset + length;
            record = new RecordSlice(type, recordOffset, payloadOffset, length);
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ushort ReadDiscoveryUInt16(int offset) {
            byte[]? bytes = _bufferedBytes;
            return bytes != null
                ? (ushort)(bytes[offset] | bytes[offset + 1] << 8)
                : ReadUInt16(_bytes, offset);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool IsCellRecordType(ushort type) =>
            type == (ushort)BiffRecordType.Blank
            || type == (ushort)BiffRecordType.BoolErr
            || type == (ushort)BiffRecordType.Formula
            || type == (ushort)BiffRecordType.Label
            || type == (ushort)BiffRecordType.LabelSst
            || type == (ushort)BiffRecordType.Number
            || type == (ushort)BiffRecordType.Rk
            || type == (ushort)BiffRecordType.MulBlank
            || type == (ushort)BiffRecordType.MulRk;

        private void ReadHeaderCells(RecordSlice record, IDictionary<int, string?> values) {
            if (record.Type == (ushort)BiffRecordType.MulRk) {
                int firstColumn = ReadUInt16(_bytes, record.PayloadOffset + 2);
                int lastColumn = ReadUInt16(_bytes, record.PayloadOffset + record.Length - 2);
                int expectedLength = checked(4 + (lastColumn - firstColumn + 1) * 6 + 2);
                if (expectedLength != record.Length) throw new InvalidDataException("A MULRK record has an invalid payload length.");
                for (int column = firstColumn; column <= lastColumn; column++) {
                    int cellOffset = record.PayloadOffset + 4 + (column - firstColumn) * 6;
                    double value = BiffRkNumberReader.ReadRkNumber(ReadUInt32(_bytes, cellOffset + 2));
                    values[column] = NumberToHeader(value, ReadUInt16(_bytes, cellOffset));
                }
                return;
            }
            if (record.Type == (ushort)BiffRecordType.MulBlank || record.Type == (ushort)BiffRecordType.Blank) return;

            int columnIndex = ReadUInt16(_bytes, record.PayloadOffset + 2);
            values[columnIndex] = ReadCellAsHeader(record);
        }

        private string? ReadCellAsHeader(RecordSlice record) {
            int payload = record.PayloadOffset;
            ushort style = record.Length >= 6 ? ReadUInt16(_bytes, payload + 4) : (ushort)0;
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.LabelSst: {
                    if (record.Length < 10) throw TruncatedCell(record);
                    uint index = ReadUInt32(_bytes, payload + 6);
                    return GetSharedString(index);
                }
                case BiffRecordType.Label:
                    return ReadLabel(record);
                case BiffRecordType.Number:
                    if (record.Length < 14) throw TruncatedCell(record);
                    return NumberToHeader(ReadDouble(_bytes, payload + 6), style);
                case BiffRecordType.Rk:
                    if (record.Length < 10) throw TruncatedCell(record);
                    return NumberToHeader(BiffRkNumberReader.ReadRkNumber(ReadUInt32(_bytes, payload + 6)), style);
                case BiffRecordType.BoolErr:
                    if (record.Length < 8) throw TruncatedCell(record);
                    return _bytes[payload + 7] != 0
                        ? BiffErrorValue.ToText(_bytes[payload + 6])
                        : (_bytes[payload + 6] != 0).ToString();
                case BiffRecordType.Formula:
                    if (record.Length < 20) throw TruncatedCell(record);
                    return ReadFormulaHeader(record, style);
                default:
                    return null;
            }
        }

        private string? ReadFormulaHeader(RecordSlice record, ushort style) {
            int result = record.PayloadOffset + 6;
            if (ReadUInt16(_bytes, result + 6) != ushort.MaxValue) {
                return NumberToHeader(ReadDouble(_bytes, result), style);
            }
            return _bytes[result] switch {
                1 => (_bytes[result + 2] != 0).ToString(),
                2 => BiffErrorValue.ToText(_bytes[result + 2]),
                3 => string.Empty,
                _ => null
            };
        }

        private void StoreCellRecord(
            RecordSlice record,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            int payload = record.PayloadOffset;
            if (record.Type == (ushort)BiffRecordType.MulRk) {
                StoreMulRk(record);
                return;
            }
            if (record.Type == (ushort)BiffRecordType.MulBlank) return;

            int column = ReadUInt16(_bytes, payload + 2);
            int ordinal = column - _firstColumn;
            if (ordinal < 0 || ordinal >= FieldCount) {
                throw new InvalidDataException($"The XLS row contains column {column} outside its discovered schema.");
            }
            ushort style = ReadUInt16(_bytes, payload + 4);
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.Blank:
                    _kinds[ordinal] = ValueKind.Empty;
                    break;
                case BiffRecordType.LabelSst:
                    if (record.Length < 10) throw TruncatedCell(record);
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = GetSharedString(ReadUInt32(_bytes, payload + 6));
                    break;
                case BiffRecordType.Label:
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = ReadLabel(record);
                    break;
                case BiffRecordType.Number:
                    if (record.Length < 14) throw TruncatedCell(record);
                    StoreNumber(ordinal, ReadDouble(_bytes, payload + 6), style);
                    break;
                case BiffRecordType.Rk:
                    if (record.Length < 10) throw TruncatedCell(record);
                    StoreNumber(ordinal, BiffRkNumberReader.ReadRkNumber(ReadUInt32(_bytes, payload + 6)), style);
                    break;
                case BiffRecordType.BoolErr:
                    if (record.Length < 8) throw TruncatedCell(record);
                    if (_bytes[payload + 7] != 0) {
                        _kinds[ordinal] = ValueKind.Error;
                        _strings[ordinal] = BiffErrorValue.ToText(_bytes[payload + 6]);
                    } else {
                        _kinds[ordinal] = ValueKind.Boolean;
                        _booleans[ordinal] = _bytes[payload + 6] != 0;
                    }
                    break;
                case BiffRecordType.Formula:
                    StoreFormula(record, ordinal, style, ref pendingFormulaOrdinal, ref pendingFormulaStyle);
                    break;
            }
        }

        private void StoreBufferedCellRecord(
            byte[] bytes,
            ushort type,
            int recordOffset,
            int payloadOffset,
            int length,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            if (type == (ushort)BiffRecordType.MulRk) {
                StoreBufferedMulRk(bytes, payloadOffset, length);
                return;
            }
            if (type == (ushort)BiffRecordType.MulBlank) return;

            if (length < 6) {
                throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
            }

            int column = bytes[payloadOffset + 2] | bytes[payloadOffset + 3] << 8;
            int ordinal = column - _firstColumn;
            ushort style = (ushort)(bytes[payloadOffset + 4] | bytes[payloadOffset + 5] << 8);
            switch ((BiffRecordType)type) {
                case BiffRecordType.Blank:
                    _kinds[ordinal] = ValueKind.Empty;
                    break;
                case BiffRecordType.LabelSst:
                    if (length < 10) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = GetSharedString(BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(payloadOffset + 6, sizeof(uint))));
                    break;
                case BiffRecordType.Label: {
                    var record = new RecordSlice(type, recordOffset, payloadOffset, length);
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = ReadLabel(record);
                    break;
                }
                case BiffRecordType.Number:
                    if (length < 14) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    StoreNumber(ordinal, ReadBufferedDouble(bytes, payloadOffset + 6), style);
                    break;
                case BiffRecordType.Rk:
                    if (length < 10) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    StoreNumber(
                        ordinal,
                        BiffRkNumberReader.ReadRkNumber(BinaryPrimitives.ReadUInt32LittleEndian(
                            bytes.AsSpan(payloadOffset + 6, sizeof(uint)))),
                        style);
                    break;
                case BiffRecordType.BoolErr:
                    if (length < 8) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    if (bytes[payloadOffset + 7] != 0) {
                        _kinds[ordinal] = ValueKind.Error;
                        _strings[ordinal] = BiffErrorValue.ToText(bytes[payloadOffset + 6]);
                    } else {
                        _kinds[ordinal] = ValueKind.Boolean;
                        _booleans[ordinal] = bytes[payloadOffset + 6] != 0;
                    }
                    break;
                case BiffRecordType.Formula:
                    StoreFormula(
                        new RecordSlice(type, recordOffset, payloadOffset, length),
                        ordinal,
                        style,
                        ref pendingFormulaOrdinal,
                        ref pendingFormulaStyle);
                    break;
            }
        }

        private void StoreMulRk(RecordSlice record) {
            int payload = record.PayloadOffset;
            int firstColumn = ReadUInt16(_bytes, payload + 2);
            int lastColumn = ReadUInt16(_bytes, payload + record.Length - 2);
            int count = checked(lastColumn - firstColumn + 1);
            int expectedLength = checked(4 + count * 6 + 2);
            if (record.Length != expectedLength) throw new InvalidDataException("A MULRK record has an invalid payload length.");
            for (int index = 0; index < count; index++) {
                int ordinal = firstColumn + index - _firstColumn;
                int cellOffset = payload + 4 + index * 6;
                StoreNumber(
                    ordinal,
                    BiffRkNumberReader.ReadRkNumber(ReadUInt32(_bytes, cellOffset + 2)),
                    ReadUInt16(_bytes, cellOffset));
            }
        }

        private void StoreBufferedMulRk(byte[] bytes, int payloadOffset, int length) {
            int firstColumn = bytes[payloadOffset + 2] | bytes[payloadOffset + 3] << 8;
            int count = (length - 6) / 6;
            int ordinal = firstColumn - _firstColumn;
            int cellOffset = payloadOffset + 4;
            for (int index = 0; index < count; index++, ordinal++, cellOffset += 6) {
                StoreNumber(
                    ordinal,
                    BiffRkNumberReader.ReadRkNumber(BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(cellOffset + 2, sizeof(uint)))),
                    BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cellOffset, sizeof(ushort))));
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static double ReadBufferedDouble(byte[] bytes, int offset) =>
            BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(
                bytes.AsSpan(offset, sizeof(long))));

        private void StoreFormula(
            RecordSlice record,
            int ordinal,
            ushort style,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            if (record.Length < 20) throw TruncatedCell(record);
            int result = record.PayloadOffset + 6;
            if (ReadUInt16(_bytes, result + 6) != ushort.MaxValue) {
                StoreNumber(ordinal, ReadDouble(_bytes, result), style);
                return;
            }
            switch (_bytes[result]) {
                case 0:
                    pendingFormulaOrdinal = ordinal;
                    pendingFormulaStyle = style;
                    break;
                case 1:
                    _kinds[ordinal] = ValueKind.Boolean;
                    _booleans[ordinal] = _bytes[result + 2] != 0;
                    break;
                case 2:
                    _kinds[ordinal] = ValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(_bytes[result + 2]);
                    break;
                case 3:
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = string.Empty;
                    break;
                default:
                    throw new InvalidDataException("A Formula record contains an invalid cached-result marker.");
            }
        }

        private void StoreFormulaString(
            RecordSlice record,
            int ordinal,
            ushort style,
            ref int nextRecordOffset) {
            _ = style;
            _kinds[ordinal] = ValueKind.Text;
            _strings[ordinal] = ReadFormulaStringValue(record, ref nextRecordOffset);
        }

        private string ReadFormulaStringValue(RecordSlice record, ref int nextRecordOffset) {
            var payloads = new List<byte[]> { CopyPayload(_bytes, record) };
            int lookahead = nextRecordOffset;
            while (TryReadRecord(_bytes, ref lookahead, out RecordSlice continuation)
                   && continuation.Type == (ushort)BiffRecordType.Continue) {
                payloads.Add(CopyPayload(_bytes, continuation));
                nextRecordOffset = lookahead;
            }
            return BiffStringReader.ReadUnicodeString(payloads);
        }

        private bool FormulaExpectsString(RecordSlice record) {
            if (record.Length < 20) throw TruncatedCell(record);
            int result = record.PayloadOffset + 6;
            return ReadUInt16(_bytes, result + 6) == ushort.MaxValue && _bytes[result] == 0;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void StoreNumber(int ordinal, double number, ushort style) {
            bool isDate = _options.TreatDatesUsingNumberFormat
                && style < _dateStyles.Length
                && _dateStyles[style]
                && LegacyXlsDateSerialConverter.TryConvert(number, _uses1904DateSystem, out _);
            _kinds[ordinal] = isDate ? ValueKind.Date : ValueKind.Number;
            _numbers[ordinal] = number;
        }

        private string ReadLabel(RecordSlice record) {
            if (record.Length < 8) throw TruncatedCell(record);
            byte[] payload = CopyPayload(_bytes, record);
            int offset = 6;
            int original = offset;
            try {
                return BiffStringReader.ReadUnicodeString(payload, ref offset);
            } catch (InvalidDataException) {
                offset = original;
                return BiffStringReader.ReadByteString(payload, ref offset);
            }
        }

        private string GetSharedString(uint index) {
            if (index >= _sharedStrings.Count) {
                throw new InvalidDataException($"XLS shared-string index {index} is outside the loaded table.");
            }
            return _sharedStrings[checked((int)index)];
        }

        private string NumberToHeader(double value, ushort style) {
            if (_options.TreatDatesUsingNumberFormat
                && style < _dateStyles.Length
                && _dateStyles[style]
                && LegacyXlsDateSerialConverter.TryConvert(value, _uses1904DateSystem, out DateTime date)) {
                return date.ToString(_options.Culture);
            }
            return value.ToString("R", _options.Culture);
        }

        private void CheckCancellation() {
            if (!CanCancelCurrentRead) return;
            if ((++_recordsSinceCancellationCheck & 1023) == 0) {
                ThrowIfReadCancellationRequested();
            }
        }

        private bool CanCancelCurrentRead =>
            _cancellationToken.CanBeCanceled || _activeReadCancellationToken.CanBeCanceled;

        private void ThrowIfReadCancellationRequested() {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_activeReadCancellationToken != _cancellationToken) {
                _activeReadCancellationToken.ThrowIfCancellationRequested();
            }
        }

        private static InvalidDataException TruncatedCell(RecordSlice record) =>
            new($"BIFF cell record 0x{record.Type:X4} at offset {record.Offset} is truncated.");

        private static InvalidDataException MissingFormulaString() =>
            new("An XLS string formula is not followed by its required String record.");

        private void ThrowIfClosed() {
            if (_closed) throw new InvalidOperationException("The XLS table reader is closed.");
        }

        private void ValidateOrdinal(int ordinal) {
            if (ordinal < 0 || ordinal >= FieldCount) {
                throw new IndexOutOfRangeException($"Column ordinal {ordinal} is outside 0..{FieldCount - 1}.");
            }
        }

        private void ValidateReadableOrdinal(int ordinal) {
            ThrowIfClosed();
            ValidateOrdinal(ordinal);
            if (!_hasCurrentRow) throw new InvalidOperationException("Read must be called before accessing values.");
        }

        private enum ValueKind : byte {
            Empty,
            Text,
            Number,
            Boolean,
            Date,
            Error
        }
    }
}
