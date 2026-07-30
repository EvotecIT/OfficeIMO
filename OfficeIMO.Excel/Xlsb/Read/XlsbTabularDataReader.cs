using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Common;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Forward-only BIFF12 worksheet reader that keeps one row of primitive values and never
    /// creates editable cells or an Open XML workbook projection.
    /// </summary>
    internal sealed partial class XlsbTabularDataReader : DbDataReader {
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
        private const int BrtBeginSheet = 129;
        private const int BrtEndSheet = 130;
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
        private readonly object?[] _customValues;
        private readonly int _firstColumn;
        private readonly int _lastDataRow;
        private readonly CancellationToken _cancellationToken;
        private bool _closed;
        private bool _hasPendingRow;
        private bool _hasCurrentRow;
        private bool _reachedEndSheetData;
        private readonly bool _hasRows;
        private int _pendingRowIndex;
        private int _nextLogicalRowIndex;
        private int _recordsSinceCancellationCheck;
        private readonly int[] _currentRowSpanBounds = new int[32];
        private int _currentRowSpanCount;

        internal XlsbTabularDataReader(
            Stream worksheetPart,
            IReadOnlyList<string> sharedStrings,
            bool[] dateStyles,
            bool uses1904DateSystem,
            bool hasHeaderRow,
            ExcelReadOptions options,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            XlsbCellReadBudget cellBudget,
            CancellationToken cancellationToken) {
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _dateStyles = dateStyles ?? throw new ArgumentNullException(nameof(dateStyles));
            _uses1904DateSystem = uses1904DateSystem;
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _limits = limits ?? throw new ArgumentNullException(nameof(limits));
            if (cellBudget == null) {
                throw new ArgumentNullException(nameof(cellBudget));
            }
            _cancellationToken = cancellationToken;
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            int actualFirstColumn = int.MaxValue;
            int actualLastColumn = -1;
            int actualFirstDataRow = -1;
            int actualLastDataRow = -1;
            try {
                DiscoverDataColumns(
                    worksheetPart,
                    limits,
                    recordBudget,
                    cellBudget,
                    cancellationToken,
                    out actualFirstColumn,
                    out actualLastColumn,
                    out actualFirstDataRow,
                    out actualLastDataRow);
            } catch {
                worksheetPart.Dispose();
                throw;
            }

            var records = new XlsbStreamRecordSliceReader(
                worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart)),
                limits.MaxRecordBytes,
                recordBudget ?? throw new ArgumentNullException(nameof(recordBudget)),
                consumeRecordBudget: false);
            _records = records;
            _lastDataRow = actualLastDataRow;
            try {
                FindSheetData(
                    actualFirstDataRow,
                    out int dimensionFirstColumn,
                    out int dimensionLastColumn);
                Dictionary<int, string?>? headerValues = null;
                if (hasHeaderRow && _hasPendingRow) {
                    int headerRowIndex = _pendingRowIndex;
                    headerValues = ReadHeaderRow();
                    if (_hasPendingRow && _pendingRowIndex <= headerRowIndex) {
                        throw new InvalidDataException(
                            $"The XLSB worksheet contains non-increasing row index {_pendingRowIndex} after header row {headerRowIndex}.");
                    }
                    _nextLogicalRowIndex = checked(headerRowIndex + 1);
                } else if (_hasPendingRow) {
                    _nextLogicalRowIndex = _pendingRowIndex;
                }

                _hasRows = _hasPendingRow;
                int headerFirstColumn = headerValues != null && headerValues.Count > 0
                    ? headerValues.Keys.Min()
                    : int.MaxValue;
                int headerLastColumn = headerValues != null && headerValues.Count > 0
                    ? headerValues.Keys.Max()
                    : -1;
                int firstColumn = int.MaxValue;
                if (dimensionLastColumn >= dimensionFirstColumn) {
                    firstColumn = dimensionFirstColumn;
                }
                if (headerFirstColumn != int.MaxValue) {
                    firstColumn = Math.Min(firstColumn, headerFirstColumn);
                }
                if (actualFirstColumn != int.MaxValue) {
                    firstColumn = Math.Min(firstColumn, actualFirstColumn);
                }
                if (firstColumn == int.MaxValue) {
                    firstColumn = 0;
                }

                int lastColumn = Math.Max(
                    Math.Max(dimensionLastColumn, headerLastColumn),
                    actualLastColumn);
                _firstColumn = firstColumn;
                int fieldCount = lastColumn >= firstColumn
                    ? checked(lastColumn - firstColumn + 1)
                    : headerValues == null || headerValues.Count == 0
                        ? 0
                        : checked(headerValues.Keys.Max() - firstColumn + 1);
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
                _customValues = new object?[fieldCount];
                _columnTypes = CreateObjectColumnTypes(fieldCount);
                _schemaRows = _options.InferSchema
                    ? BufferSchemaRows()
                    : null;
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
            if (_schemaRows != null && _schemaRowIndex < _schemaRows.Count) {
                LoadBufferedRow(_schemaRows[_schemaRowIndex++]);
                _hasCurrentRow = true;
                return true;
            }

            return ReadSourceRow();
        }

        private bool ReadSourceRow() {
            if (!_hasPendingRow) {
                if (_reachedEndSheetData) {
                    return false;
                }

                throw new InvalidDataException(
                    "The XLSB worksheet ended before the required BrtEndSheetData record.");
            }

            Array.Clear(_kinds, 0, _kinds.Length);
            Array.Clear(_strings, 0, _strings.Length);
            Array.Clear(_customValues, 0, _customValues.Length);
            if (_nextLogicalRowIndex < _pendingRowIndex) {
                _nextLogicalRowIndex++;
                _hasCurrentRow = true;
                return true;
            }

            int currentRowIndex = _pendingRowIndex;
            _hasPendingRow = false;
            bool reachedRowBoundary = false;
            if (_cancellationToken.CanBeCanceled) {
                while (_records.TryRead(out XlsbRecordSlice record)) {
                    CheckCancellation();
                    if (record.Type == BrtRowHdr) {
                        if (TrySetPendingRow(record)) {
                            reachedRowBoundary = true;
                            break;
                        }

                        continue;
                    }

                    if (record.Type == BrtEndSheetData) {
                        _reachedEndSheetData = true;
                        reachedRowBoundary = true;
                        break;
                    }

                    if (IsCellRecord(record.Type)) {
                        StoreCell(record);
                    }
                }
            } else {
                while (_records.TryRead(out XlsbRecordSlice record)) {
                    if (record.Type == BrtRowHdr) {
                        if (TrySetPendingRow(record)) {
                            reachedRowBoundary = true;
                            break;
                        }

                        continue;
                    }

                    if (record.Type == BrtEndSheetData) {
                        _reachedEndSheetData = true;
                        reachedRowBoundary = true;
                        break;
                    }

                    if (IsCellRecord(record.Type)) {
                        StoreCell(record);
                    }
                }
            }

            if (!reachedRowBoundary) {
                throw new InvalidDataException(
                    "The XLSB worksheet ended before the required BrtEndSheetData record.");
            }

            _nextLogicalRowIndex = checked(currentRowIndex + 1);
            if (_hasPendingRow && _pendingRowIndex <= currentRowIndex) {
                throw new InvalidDataException(
                    $"The XLSB worksheet contains non-increasing row index {_pendingRowIndex} after row {currentRowIndex}.");
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
            _schemaRows?.Clear();
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Close();
            }

            base.Dispose(disposing);
        }

        private void FindSheetData(
            int firstDataRow,
            out int firstColumn,
            out int lastColumn) {
            firstColumn = 0;
            lastColumn = -1;
            bool inSheetData = false;
            bool sawDimension = false;
            while (_records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtWsDim) {
                    if (sawDimension) {
                        throw new InvalidDataException(
                            "The XLSB worksheet contains more than one BrtWsDim record.");
                    }
                    if (record.Size != 16) {
                        throw new InvalidDataException(
                            $"The BrtWsDim record at offset {record.RecordOffset} has invalid payload length {record.Size}.");
                    }
                    sawDimension = true;
                    var cursor = record.CreateCursor();
                    uint firstRow = cursor.ReadUInt32();
                    uint lastRow = cursor.ReadUInt32();
                    uint firstColumnValue = cursor.ReadUInt32();
                    uint lastColumnValue = cursor.ReadUInt32();
                    if (firstRow > lastRow
                        || lastRow >= A1.MaxRows
                        || firstColumnValue > lastColumnValue
                        || lastColumnValue >= A1.MaxColumns) {
                        throw new InvalidDataException(
                            $"The BrtWsDim record at offset {record.RecordOffset} contains an invalid worksheet range.");
                    }
                    firstColumn = checked((int)firstColumnValue);
                    lastColumn = checked((int)lastColumnValue);
                } else if (record.Type == BrtBeginSheetData) {
                    inSheetData = true;
                } else if (inSheetData && record.Type == BrtRowHdr) {
                    int rowIndex = ValidateRowHeader(
                        record,
                        _currentRowSpanBounds,
                        out _currentRowSpanCount);
                    if (rowIndex == firstDataRow) {
                        _pendingRowIndex = rowIndex;
                        _hasPendingRow = true;
                        return;
                    }
                } else if (inSheetData && record.Type == BrtEndSheetData) {
                    _reachedEndSheetData = true;
                    return;
                }
            }

            if (inSheetData) {
                throw new InvalidDataException(
                    "The XLSB worksheet ended before the required BrtEndSheetData record.");
            }
        }

        private Dictionary<int, string?> ReadHeaderRow() {
            var values = new Dictionary<int, string?>();
            _hasPendingRow = false;
            while (_records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtRowHdr) {
                    if (TrySetPendingRow(record)) {
                        break;
                    }

                    continue;
                }

                if (record.Type == BrtEndSheetData) {
                    _reachedEndSheetData = true;
                    break;
                }

                if (!IsCellRecord(record.Type)) {
                    continue;
                }

                DecodedCell cell = DecodeCell(record);
                values[cell.Column] = CellToHeaderText(cell);
            }

            if (!_hasPendingRow && !_reachedEndSheetData) {
                throw new InvalidDataException(
                    "The XLSB worksheet ended before the required BrtEndSheetData record.");
            }

            return values;
        }

        private void StoreCell(XlsbRecordSlice record) {
            if (_options.CellValueConverter == null) {
                StoreCellFast(record);
                return;
            }

            DecodedCell cell = DecodeCell(record);
            int ordinal = cell.Column - _firstColumn;
            if (ordinal < 0 || ordinal >= FieldCount) {
                throw new InvalidDataException(
                    $"The XLSB row contains column {cell.Column} outside the schema established by its header or worksheet dimension.");
            }

            _kinds[ordinal] = cell.Kind;
            switch (cell.Kind) {
                case XlsbTabularValueKind.Text:
                case XlsbTabularValueKind.Error:
                    _strings[ordinal] = cell.Text;
                    break;
                case XlsbTabularValueKind.Number:
                case XlsbTabularValueKind.Date:
                    _numbers[ordinal] = cell.Number;
                    break;
                case XlsbTabularValueKind.Boolean:
                    _booleans[ordinal] = cell.Boolean;
                    break;
                case XlsbTabularValueKind.Custom:
                    _customValues[ordinal] = cell.CustomValue;
                    break;
            }
        }

        private DecodedCell DecodeCell(XlsbRecordSlice record) {
            EnsureFormulaModeSupported(record.Type);
            var cursor = record.CreateCursor();
            int column = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (column < 0 || column >= A1.MaxColumns) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
            }
            ValidateCellCoveredByCurrentRow(column, record);
            ValidateStyleIndex(styleIndex, record);

            bool isDate = _options.TreatDatesUsingNumberFormat
                && styleIndex < _dateStyles.Length
                && _dateStyles[styleIndex];
            DecodedCell cell;
            switch (record.Type) {
                case BrtCellBlank:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Empty);
                    break;
                case BrtCellRk:
                    cell = NumericCell(column, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), isDate);
                    break;
                case BrtCellError:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Error) {
                        Text = BiffErrorValue.ToText(cursor.ReadByte())
                    };
                    break;
                case BrtCellBool:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Boolean) {
                        Boolean = cursor.ReadByte() != 0
                    };
                    break;
                case BrtCellReal:
                    cell = NumericCell(column, cursor.ReadDouble(), isDate);
                    break;
                case BrtCellSt:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                    break;
                case BrtCellIsst: {
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= _sharedStrings.Count) {
                        throw new InvalidDataException(
                            $"The XLSB cell refers to missing shared string {sharedStringIndex}.");
                    }

                    cell = new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = _sharedStrings[checked((int)sharedStringIndex)],
                        RawText = sharedStringIndex.ToString(CultureInfo.InvariantCulture)
                    };
                    break;
                }
                case BrtCellRString:
                    cursor.ReadByte();
                    cell = new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                    break;
                case BrtFmlaString:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaNum:
                    cell = NumericCell(column, cursor.ReadDouble(), isDate);
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaBool:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Boolean) {
                        Boolean = cursor.ReadByte() != 0
                    };
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaError:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Error) {
                        Text = BiffErrorValue.ToText(cursor.ReadByte())
                    };
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }

            return ApplyCellValueConverter(cell, styleIndex, record.Type);
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
                XlsbTabularValueKind.Custom => Convert.ToString(cell.CustomValue, _options.Culture),
                _ => null
            };

        private DecodedCell ApplyCellValueConverter(
            DecodedCell cell,
            uint styleIndex,
            int recordType) {
            Func<ExcelCellContext, ExcelCellValue>? converter = _options.CellValueConverter;
            if (converter == null) {
                return cell;
            }

            CellValues? typeHint = cell.Kind switch {
                XlsbTabularValueKind.Number or XlsbTabularValueKind.Date => CellValues.Number,
                XlsbTabularValueKind.Boolean => CellValues.Boolean,
                XlsbTabularValueKind.Error => CellValues.Error,
                XlsbTabularValueKind.Text when recordType == BrtCellIsst => CellValues.SharedString,
                XlsbTabularValueKind.Text => CellValues.String,
                _ => null
            };
            string? rawText = cell.RawText ?? cell.Kind switch {
                XlsbTabularValueKind.Number or XlsbTabularValueKind.Date =>
                    cell.Number.ToString("R", CultureInfo.InvariantCulture),
                XlsbTabularValueKind.Boolean => cell.Boolean ? "1" : "0",
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => cell.Text,
                _ => null
            };
            string? inlineText = cell.Kind == XlsbTabularValueKind.Text
                                 && recordType != BrtCellIsst
                ? cell.Text
                : null;
            ExcelCellValue converted = converter(
                new ExcelCellContext(
                    typeHint,
                    styleIndex,
                    rawText,
                    inlineText,
                    _options.Culture));
            return converted.Handled
                ? new DecodedCell(cell.Column, XlsbTabularValueKind.Custom) {
                    CustomValue = converted.Value
                }
                : cell;
        }

        private DateTime ConvertDate(double serial) {
            if (LegacyXlsDateSerialConverter.TryConvert(serial, _uses1904DateSystem, out DateTime value)) {
                return value;
            }

            throw new InvalidCastException($"The XLSB numeric value '{serial}' is not a valid Excel date.");
        }

        private static decimal ConvertExcelNumberToDecimal(double number) {
            if (TryConvertExcelNumberToDecimal(number, out decimal value)) {
                return value;
            }

            throw new InvalidCastException(
                $"The XLSB numeric value '{number}' cannot be represented as decimal.");
        }

        private static bool TryConvertExcelNumberToDecimal(double number, out decimal value) {
            if (double.IsNaN(number) || double.IsInfinity(number)) {
                value = default;
                return false;
            }

            try {
                value = (decimal)number;
                return true;
            } catch (OverflowException) {
                value = default;
                return false;
            }
        }

        private static bool IsCellRecord(int recordType) =>
            recordType is >= BrtCellBlank and <= BrtFmlaError or BrtCellRString;

        private static bool IsFormulaRecord(int recordType) =>
            recordType is >= BrtFmlaString and <= BrtFmlaError;

        private void EnsureFormulaModeSupported(int recordType) {
            if (!_options.UseCachedFormulaResult && IsFormulaRecord(recordType)) {
                throw new NotSupportedException(
                    "XLSB formula-token projection is not supported when cached formula results are disabled.");
            }
        }

        private bool TrySetPendingRow(XlsbRecordSlice record) {
            int rowIndex = ValidateRowHeader(
                record,
                _currentRowSpanBounds,
                out _currentRowSpanCount);
            if (rowIndex > _lastDataRow) {
                return false;
            }

            _pendingRowIndex = rowIndex;
            _hasPendingRow = true;
            return true;
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

        private static int ValidateRowHeader(
            XlsbRecordSlice record,
            int[] spanBounds,
            out int spanCount) {
            if (record.Size < 17) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} is truncated.");
            }
            if (spanBounds == null || spanBounds.Length < 32) {
                throw new ArgumentException(
                    "A row-span buffer must hold all 16 BIFF12 column spans.",
                    nameof(spanBounds));
            }

            var cursor = record.CreateCursor();
            uint rowIndex = cursor.ReadUInt32();
            cursor.ReadUInt32();
            cursor.ReadUInt16();
            byte extraFlags = cursor.ReadByte();
            byte flags = cursor.ReadByte();
            byte phoneticFlags = cursor.ReadByte();
            uint declaredSpanCount = cursor.ReadUInt32();
            if (rowIndex >= A1.MaxRows
                || (extraFlags & 0xFC) != 0
                || (flags & 0x80) != 0
                || (phoneticFlags & 0xFE) != 0
                || declaredSpanCount > 16) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} contains invalid row metadata.");
            }
            if (cursor.Remaining != checked((int)declaredSpanCount * 8)) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} has an invalid column-span payload.");
            }

            int previousLast = -1;
            spanCount = checked((int)declaredSpanCount);
            for (int index = 0; index < spanCount; index++) {
                uint firstColumn = cursor.ReadUInt32();
                uint lastColumn = cursor.ReadUInt32();
                if (firstColumn > lastColumn
                    || lastColumn >= A1.MaxColumns
                    || firstColumn / 1024U != lastColumn / 1024U
                    || firstColumn <= previousLast) {
                    throw new InvalidDataException(
                        $"The BrtRowHdr record at offset {record.RecordOffset} contains an invalid column span.");
                }
                int offset = index * 2;
                spanBounds[offset] = checked((int)firstColumn);
                spanBounds[offset + 1] = checked((int)lastColumn);
                previousLast = checked((int)lastColumn);
            }

            return checked((int)rowIndex);
        }

        private void ValidateCellCoveredByCurrentRow(
            int column,
            XlsbRecordSlice record) {
            for (int index = 0; index < _currentRowSpanCount; index++) {
                int offset = index * 2;
                if (_currentRowSpanBounds[offset] <= column
                    && column <= _currentRowSpanBounds[offset + 1]) {
                    return;
                }
            }

            throw new InvalidDataException(
                $"The XLSB cell record at offset {record.RecordOffset} for column {column} is not covered by its BrtRowHdr column spans.");
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
                return source.Length;
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

        private Type GetValueType(int ordinal) =>
            _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => typeof(string),
                XlsbTabularValueKind.Number => _options.NumericAsDecimal
                    && TryConvertExcelNumberToDecimal(_numbers[ordinal], out _)
                        ? typeof(decimal)
                        : typeof(double),
                XlsbTabularValueKind.Boolean => typeof(bool),
                XlsbTabularValueKind.Date => typeof(DateTime),
                XlsbTabularValueKind.Custom => _customValues[ordinal]?.GetType() ?? typeof(object),
                _ => typeof(object)
            };

        private object GetNumericValue(double number) {
            if (_options.NumericAsDecimal && TryConvertExcelNumberToDecimal(number, out decimal value)) {
                return value;
            }

            return number;
        }

        private enum XlsbTabularValueKind : byte {
            Empty,
            Text,
            Number,
            Boolean,
            Date,
            Error,
            Custom
        }

        private struct DecodedCell {
            internal DecodedCell(int column, XlsbTabularValueKind kind) {
                Column = column;
                Kind = kind;
                Number = 0;
                Boolean = false;
                Text = null;
                RawText = null;
                CustomValue = null;
            }

            internal int Column { get; }

            internal XlsbTabularValueKind Kind { get; }

            internal double Number { get; set; }

            internal bool Boolean { get; set; }

            internal string? Text { get; set; }

            internal string? RawText { get; set; }

            internal object? CustomValue { get; set; }
        }
    }
}
