using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Forward-only BIFF12 worksheet reader that keeps one row of primitive values and never
    /// creates editable cells or an Open XML workbook projection.
    /// </summary>
    internal sealed class XlsbTabularDataReader : DbDataReader {
        private const int BrtRowHdr = 0;
        private const int BrtCellBlank = 1;
        private const int BrtCellRk = 2;
        private const int BrtCellError = 3;
        private const int BrtCellBool = 4;
        private const int BrtCellReal = 5;
        private const int BrtCellSt = 6;
        private const int BrtCellIsst = 7;
        private const int BrtFmlaString = 8;
        private const int BrtFmlaNum = 9;
        private const int BrtFmlaBool = 10;
        private const int BrtFmlaError = 11;
        private const int BrtCellRString = 62;
        private const int BrtBeginSheetData = 145;
        private const int BrtEndSheetData = 146;
        private const int BrtWsDim = 148;

        private readonly XlsbStreamRecordSliceReader _records;
        private readonly IReadOnlyList<string> _sharedStrings;
        private readonly bool[] _dateStyles;
        private readonly bool _uses1904DateSystem;
        private readonly ExcelReadOptions _options;
        private readonly XlsbImportOptions _limits;
        private readonly string[] _headers;
        private readonly Dictionary<string, int> _ordinals;
        private readonly XlsbTabularValueKind[] _kinds;
        private readonly double[] _numbers;
        private readonly bool[] _booleans;
        private readonly string?[] _strings;
        private readonly int _firstColumn;
        private readonly CancellationToken _cancellationToken;
        private bool _closed;
        private bool _hasPendingRow;
        private bool _hasCurrentRow;
        private readonly bool _hasRows;
        private int _cellsRead;
        private int _recordsSinceCancellationCheck;

        internal XlsbTabularDataReader(
            Stream worksheetPart,
            IReadOnlyList<string> sharedStrings,
            bool[] dateStyles,
            bool uses1904DateSystem,
            bool hasHeaderRow,
            ExcelReadOptions options,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            CancellationToken cancellationToken) {
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _dateStyles = dateStyles ?? throw new ArgumentNullException(nameof(dateStyles));
            _uses1904DateSystem = uses1904DateSystem;
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _limits = limits ?? throw new ArgumentNullException(nameof(limits));
            _cancellationToken = cancellationToken;
            var records = new XlsbStreamRecordSliceReader(
                worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart)),
                limits.MaxRecordBytes,
                recordBudget ?? throw new ArgumentNullException(nameof(recordBudget)));
            _records = records;
            try {
                FindSheetData(out int dimensionFirstColumn, out int dimensionLastColumn);
                _firstColumn = dimensionFirstColumn;
                Dictionary<int, string?>? headerValues = null;
                if (hasHeaderRow && _hasPendingRow) {
                    headerValues = ReadHeaderRow();
                }

                _hasRows = _hasPendingRow;
                int fieldCount = dimensionLastColumn >= dimensionFirstColumn
                    ? checked(dimensionLastColumn - dimensionFirstColumn + 1)
                    : headerValues == null || headerValues.Count == 0
                        ? 0
                        : checked(headerValues.Keys.Max() - dimensionFirstColumn + 1);
                if (fieldCount > _options.MaxDataReaderColumns) {
                    throw new InvalidDataException(
                        $"XLSB table column count {fieldCount} exceeds the configured limit of {_options.MaxDataReaderColumns}.");
                }

                _headers = ExcelHeaderNameHelper.BuildUniqueHeaders(
                    fieldCount,
                    ordinal => headerValues != null
                        && headerValues.TryGetValue(ordinal + _firstColumn, out string? value)
                            ? value
                            : null,
                    _options.NormalizeHeaders);
                _ordinals = new Dictionary<string, int>(_headers.Length, StringComparer.OrdinalIgnoreCase);
                for (int index = 0; index < _headers.Length; index++) {
                    _ordinals[_headers[index]] = index;
                }

                _kinds = new XlsbTabularValueKind[fieldCount];
                _numbers = new double[fieldCount];
                _booleans = new bool[fieldCount];
                _strings = new string?[fieldCount];
            } catch {
                records.Dispose();
                throw;
            }
        }

        public override object this[int ordinal] => GetValue(ordinal);

        public override object this[string name] => GetValue(GetOrdinal(name));

        public override int Depth => 0;

        public override int FieldCount => _headers.Length;

        public override bool HasRows => _hasRows;

        public override bool IsClosed => _closed;

        public override int RecordsAffected => -1;

        public override bool Read() {
            ThrowIfClosed();
            _cancellationToken.ThrowIfCancellationRequested();
            _hasCurrentRow = false;
            if (!_hasPendingRow) {
                return false;
            }

            Array.Clear(_kinds, 0, _kinds.Length);
            Array.Clear(_strings, 0, _strings.Length);
            _hasPendingRow = false;
            while (_records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtRowHdr) {
                    SetPendingRow(record);
                    break;
                }

                if (record.Type == BrtEndSheetData) {
                    break;
                }

                if (IsCellRecord(record.Type)) {
                    StoreCell(record);
                }
            }

            _hasCurrentRow = true;
            return true;
        }

        public override bool NextResult() => false;

        public override void Close() {
            if (_closed) {
                return;
            }

            _closed = true;
            _records.Dispose();
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Close();
            }

            base.Dispose(disposing);
        }

        public override string GetName(int ordinal) {
            ValidateOrdinal(ordinal);
            return _headers[ordinal];
        }

        public override int GetOrdinal(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            if (_ordinals.TryGetValue(name, out int ordinal)) {
                return ordinal;
            }

            throw new IndexOutOfRangeException($"Column '{name}' was not found.");
        }

        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
        public override Type GetFieldType(int ordinal) {
            ValidateOrdinal(ordinal);
            if (!_hasCurrentRow) {
                return typeof(object);
            }

            return _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => typeof(string),
                XlsbTabularValueKind.Number => _options.NumericAsDecimal ? typeof(decimal) : typeof(double),
                XlsbTabularValueKind.Boolean => typeof(bool),
                XlsbTabularValueKind.Date => typeof(DateTime),
                _ => typeof(object)
            };
        }

        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

        public override object GetValue(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => _strings[ordinal]!,
                XlsbTabularValueKind.Number => _options.NumericAsDecimal
                    ? (object)ConvertExcelNumberToDecimal(_numbers[ordinal])
                    : _numbers[ordinal],
                XlsbTabularValueKind.Boolean => _booleans[ordinal],
                XlsbTabularValueKind.Date => ConvertDate(_numbers[ordinal]),
                _ => DBNull.Value
            };
        }

        public override int GetValues(object[] values) {
            if (values == null) {
                throw new ArgumentNullException(nameof(values));
            }

            int count = Math.Min(values.Length, FieldCount);
            for (int index = 0; index < count; index++) {
                values[index] = GetValue(index);
            }

            return count;
        }

        public override bool IsDBNull(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Empty;
        }

        public override string GetString(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => _strings[ordinal]!,
                XlsbTabularValueKind.Number => _numbers[ordinal].ToString("R", _options.Culture),
                XlsbTabularValueKind.Boolean => _booleans[ordinal].ToString(),
                XlsbTabularValueKind.Date => ConvertDate(_numbers[ordinal]).ToString(_options.Culture),
                _ => throw new InvalidCastException("The XLSB cell is blank.")
            };
        }

        public override bool GetBoolean(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            if (_kinds[ordinal] == XlsbTabularValueKind.Boolean) {
                return _booleans[ordinal];
            }

            return Convert.ToBoolean(GetValue(ordinal), _options.Culture);
        }

        public override byte GetByte(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToByte(_numbers[ordinal])
                : Convert.ToByte(GetValue(ordinal), _options.Culture);
        }

        public override short GetInt16(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt16(_numbers[ordinal])
                : Convert.ToInt16(GetValue(ordinal), _options.Culture);
        }

        public override int GetInt32(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt32(_numbers[ordinal])
                : Convert.ToInt32(GetValue(ordinal), _options.Culture);
        }

        public override long GetInt64(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt64(_numbers[ordinal])
                : Convert.ToInt64(GetValue(ordinal), _options.Culture);
        }

        public override float GetFloat(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? (float)_numbers[ordinal]
                : Convert.ToSingle(GetValue(ordinal), _options.Culture);
        }

        public override double GetDouble(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? _numbers[ordinal]
                : Convert.ToDouble(GetValue(ordinal), _options.Culture);
        }

        public override decimal GetDecimal(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? ConvertExcelNumberToDecimal(_numbers[ordinal])
                : Convert.ToDecimal(GetValue(ordinal), _options.Culture);
        }

        public override DateTime GetDateTime(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Date
                ? ConvertDate(_numbers[ordinal])
                : Convert.ToDateTime(GetValue(ordinal), _options.Culture);
        }

        public override Guid GetGuid(int ordinal) {
            object value = GetValue(ordinal);
            return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _options.Culture)!);
        }

        public override char GetChar(int ordinal) => Convert.ToChar(GetValue(ordinal), _options.Culture);

        public override long GetBytes(
            int ordinal,
            long dataOffset,
            byte[]? buffer,
            int bufferOffset,
            int length) {
            if (GetValue(ordinal) is not byte[] bytes) {
                throw new InvalidCastException("The XLSB cell does not contain binary data.");
            }

            return CopySegment(bytes, dataOffset, buffer, bufferOffset, length);
        }

        public override long GetChars(
            int ordinal,
            long dataOffset,
            char[]? buffer,
            int bufferOffset,
            int length) {
            char[] characters = GetString(ordinal).ToCharArray();
            return CopySegment(characters, dataOffset, buffer, bufferOffset, length);
        }

        public override IEnumerator GetEnumerator() => new DbEnumerator(this, closeReader: false);

        [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type.TypeInitializer or other Type members.")]
        public override DataTable GetSchemaTable() {
            var table = new DataTable("SchemaTable");
            table.Columns.Add("ColumnName", typeof(string));
            table.Columns.Add("ColumnOrdinal", typeof(int));
            table.Columns.Add("DataType", typeof(Type));
            for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                DataRow row = table.NewRow();
                row["ColumnName"] = GetName(ordinal);
                row["ColumnOrdinal"] = ordinal;
                row["DataType"] = GetFieldType(ordinal);
                table.Rows.Add(row);
            }

            return table;
        }

        private void FindSheetData(out int firstColumn, out int lastColumn) {
            firstColumn = 0;
            lastColumn = -1;
            bool inSheetData = false;
            while (_records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtWsDim) {
                    var cursor = record.CreateCursor();
                    cursor.ReadUInt32();
                    cursor.ReadUInt32();
                    firstColumn = checked((int)cursor.ReadUInt32());
                    lastColumn = checked((int)cursor.ReadUInt32());
                } else if (record.Type == BrtBeginSheetData) {
                    inSheetData = true;
                } else if (inSheetData && record.Type == BrtRowHdr) {
                    SetPendingRow(record);
                    return;
                } else if (inSheetData && record.Type == BrtEndSheetData) {
                    return;
                }
            }
        }

        private Dictionary<int, string?> ReadHeaderRow() {
            var values = new Dictionary<int, string?>();
            _hasPendingRow = false;
            while (_records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtRowHdr) {
                    SetPendingRow(record);
                    break;
                }

                if (record.Type == BrtEndSheetData) {
                    break;
                }

                if (!IsCellRecord(record.Type)) {
                    continue;
                }

                DecodedCell cell = DecodeCell(record);
                values[cell.Column] = CellToHeaderText(cell);
            }

            return values;
        }

        private void StoreCell(XlsbRecordSlice record) {
            var cursor = record.CreateCursor();
            int column = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (column < 0 || column >= A1.MaxColumns) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
            }

            int ordinal = column - _firstColumn;
            if (ordinal < 0 || ordinal >= FieldCount) {
                return;
            }

            _cellsRead = checked(_cellsRead + 1);
            if (_cellsRead > _limits.MaxCells) {
                throw new InvalidDataException(
                    $"The XLSB table exceeds the configured limit of {_limits.MaxCells} populated cells.");
            }

            bool isDate = _options.TreatDatesUsingNumberFormat
                && styleIndex < _dateStyles.Length
                && _dateStyles[styleIndex];
            switch (record.Type) {
                case BrtCellBlank:
                    _kinds[ordinal] = XlsbTabularValueKind.Empty;
                    break;
                case BrtCellRk:
                    StoreNumber(ordinal, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), isDate);
                    break;
                case BrtCellError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(cursor.ReadByte());
                    break;
                case BrtCellBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = cursor.ReadByte() != 0;
                    break;
                case BrtCellReal:
                    StoreNumber(ordinal, cursor.ReadDouble(), isDate);
                    break;
                case BrtCellSt:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    break;
                case BrtCellIsst: {
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= _sharedStrings.Count) {
                        throw new InvalidDataException(
                            $"The XLSB cell refers to missing shared string {sharedStringIndex}.");
                    }

                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = _sharedStrings[checked((int)sharedStringIndex)];
                    break;
                }
                case BrtCellRString:
                    cursor.ReadByte();
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    break;
                case BrtFmlaString:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    break;
                case BrtFmlaNum:
                    StoreNumber(ordinal, cursor.ReadDouble(), isDate);
                    break;
                case BrtFmlaBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = cursor.ReadByte() != 0;
                    break;
                case BrtFmlaError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(cursor.ReadByte());
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }
        }

        private void StoreNumber(int ordinal, double number, bool isDate) {
            _kinds[ordinal] = isDate ? XlsbTabularValueKind.Date : XlsbTabularValueKind.Number;
            _numbers[ordinal] = number;
        }

        private DecodedCell DecodeCell(XlsbRecordSlice record) {
            var cursor = record.CreateCursor();
            int column = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (column < 0 || column >= A1.MaxColumns) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
            }

            bool isDate = _options.TreatDatesUsingNumberFormat
                && styleIndex < _dateStyles.Length
                && _dateStyles[styleIndex];
            switch (record.Type) {
                case BrtCellBlank:
                    return new DecodedCell(column, XlsbTabularValueKind.Empty);
                case BrtCellRk:
                    return NumericCell(column, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), isDate);
                case BrtCellError:
                    return new DecodedCell(column, XlsbTabularValueKind.Error) {
                        Text = BiffErrorValue.ToText(cursor.ReadByte())
                    };
                case BrtCellBool:
                    return new DecodedCell(column, XlsbTabularValueKind.Boolean) {
                        Boolean = cursor.ReadByte() != 0
                    };
                case BrtCellReal:
                    return NumericCell(column, cursor.ReadDouble(), isDate);
                case BrtCellSt:
                    return new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                case BrtCellIsst: {
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= _sharedStrings.Count) {
                        throw new InvalidDataException(
                            $"The XLSB cell refers to missing shared string {sharedStringIndex}.");
                    }

                    return new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = _sharedStrings[checked((int)sharedStringIndex)]
                    };
                }
                case BrtCellRString:
                    cursor.ReadByte();
                    return new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                case BrtFmlaString:
                    return new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                case BrtFmlaNum:
                    return NumericCell(column, cursor.ReadDouble(), isDate);
                case BrtFmlaBool:
                    return new DecodedCell(column, XlsbTabularValueKind.Boolean) {
                        Boolean = cursor.ReadByte() != 0
                    };
                case BrtFmlaError:
                    return new DecodedCell(column, XlsbTabularValueKind.Error) {
                        Text = BiffErrorValue.ToText(cursor.ReadByte())
                    };
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }
        }

        private static DecodedCell NumericCell(int column, double number, bool isDate) =>
            new(column, isDate ? XlsbTabularValueKind.Date : XlsbTabularValueKind.Number) {
                Number = number
            };

        private string? CellToHeaderText(DecodedCell cell) =>
            cell.Kind switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => cell.Text,
                XlsbTabularValueKind.Number => cell.Number.ToString("R", _options.Culture),
                XlsbTabularValueKind.Boolean => cell.Boolean.ToString(),
                XlsbTabularValueKind.Date => ConvertDate(cell.Number).ToString(_options.Culture),
                _ => null
            };

        private DateTime ConvertDate(double serial) {
            if (LegacyXlsDateSerialConverter.TryConvert(serial, _uses1904DateSystem, out DateTime value)) {
                return value;
            }

            throw new InvalidCastException($"The XLSB numeric value '{serial}' is not a valid Excel date.");
        }

        private static decimal ConvertExcelNumberToDecimal(double number) {
            if (double.IsNaN(number) || double.IsInfinity(number)) {
                throw new InvalidCastException($"The XLSB numeric value '{number}' cannot be represented as decimal.");
            }

            try {
                return (decimal)number;
            } catch (OverflowException exception) {
                throw new InvalidCastException(
                    $"The XLSB numeric value '{number}' cannot be represented as decimal.",
                    exception);
            }
        }

        private static bool IsCellRecord(int recordType) =>
            recordType is >= BrtCellBlank and <= BrtFmlaError or BrtCellRString;

        private void SetPendingRow(XlsbRecordSlice record) {
            ValidateRowHeader(record);
            _hasPendingRow = true;
        }

        private void CheckCancellation() {
            if (!_cancellationToken.CanBeCanceled) {
                return;
            }

            _recordsSinceCancellationCheck++;
            if ((_recordsSinceCancellationCheck & 1023) == 0) {
                _cancellationToken.ThrowIfCancellationRequested();
            }
        }

        private static void ValidateRowHeader(XlsbRecordSlice record) {
            if (record.Size < 17) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} is truncated.");
            }

            var cursor = record.CreateCursor();
            uint rowIndex = cursor.ReadUInt32();
            if (rowIndex >= A1.MaxRows) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} contains invalid row index {rowIndex}.");
            }
        }

        private static long CopySegment<T>(
            T[] source,
            long dataOffset,
            T[]? destination,
            int destinationOffset,
            int length) {
            if (dataOffset < 0 || dataOffset > source.Length) {
                throw new ArgumentOutOfRangeException(nameof(dataOffset));
            }

            int available = source.Length - checked((int)dataOffset);
            if (destination == null) {
                return available;
            }

            int count = Math.Min(available, length);
            Array.Copy(source, checked((int)dataOffset), destination, destinationOffset, count);
            return count;
        }

        private void ValidateOrdinal(int ordinal) {
            if (ordinal < 0 || ordinal >= FieldCount) {
                throw new IndexOutOfRangeException($"Column ordinal {ordinal} is outside 0..{FieldCount - 1}.");
            }
        }

        private void ValidateReadableOrdinal(int ordinal) {
            ThrowIfClosed();
            ValidateOrdinal(ordinal);
            if (!_hasCurrentRow) {
                throw new InvalidOperationException("Read must be called before accessing values.");
            }
        }

        private void ThrowIfClosed() {
            if (_closed) {
                throw new InvalidOperationException("The XLSB table reader is closed.");
            }
        }

        private enum XlsbTabularValueKind : byte {
            Empty,
            Text,
            Number,
            Boolean,
            Date,
            Error
        }

        private struct DecodedCell {
            internal DecodedCell(int column, XlsbTabularValueKind kind) {
                Column = column;
                Kind = kind;
                Number = 0;
                Boolean = false;
                Text = null;
            }

            internal int Column { get; }

            internal XlsbTabularValueKind Kind { get; }

            internal double Number { get; set; }

            internal bool Boolean { get; set; }

            internal string? Text { get; set; }
        }
    }
}
