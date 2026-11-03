using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

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

        private readonly XlsbRecordSliceReader _records;
        private readonly XlsbPooledPartStream _worksheetPart;
        private readonly IReadOnlyList<string> _sharedStrings;
        private readonly ExcelSerialDateStyle[] _dateStyles;
        private readonly bool _uses1904DateSystem;
        private readonly ExcelReaderOptions _options;
        private readonly XlsbImportOptions _limits;
        private readonly string[] _headers;
        private readonly Dictionary<string, int> _ordinals;
        private readonly XlsbTabularValueKind[] _kinds;
        private readonly double[] _numbers;
        private readonly bool[] _booleans;
        private readonly string?[] _strings;
        private readonly object?[] _customValues;
        private readonly int _firstColumn;
        private readonly CancellationToken _cancellationToken;
        private CancellationToken _activeReadCancellationToken;
        private readonly XlsbRecordReadBudget _recordBudget;
        private readonly XlsbCellReadBudget _cellBudget;
        private readonly XlsbLogicalRowReadBudget _logicalRowBudget;
        private bool _closed;
        private bool _hasPendingRow;
        private bool _hasCurrentRow;
        private readonly bool _hasRows;
        private int _pendingRowIndex;
        private int _nextLogicalRowIndex;
        private int _recordsSinceCancellationCheck;

        internal XlsbTabularDataReader(
            Stream worksheetPart,
            IReadOnlyList<string> sharedStrings,
            ExcelSerialDateStyle[] dateStyles,
            bool uses1904DateSystem,
            bool hasHeaderRow,
            ExcelReaderOptions options,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            XlsbCellReadBudget cellBudget,
            XlsbLogicalRowReadBudget logicalRowBudget,
            CancellationToken cancellationToken)
            : this(
                CreatePooledPart(worksheetPart, limits, cancellationToken),
                sharedStrings,
                dateStyles,
                uses1904DateSystem,
                hasHeaderRow,
                options,
                limits,
                recordBudget,
                cellBudget,
                logicalRowBudget,
                cancellationToken) {
        }

        internal XlsbTabularDataReader(
            XlsbPooledPartStream worksheetPart,
            IReadOnlyList<string> sharedStrings,
            ExcelSerialDateStyle[] dateStyles,
            bool uses1904DateSystem,
            bool hasHeaderRow,
            ExcelReaderOptions options,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            XlsbCellReadBudget cellBudget,
            XlsbLogicalRowReadBudget logicalRowBudget,
            CancellationToken cancellationToken) {
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _dateStyles = dateStyles ?? throw new ArgumentNullException(nameof(dateStyles));
            _uses1904DateSystem = uses1904DateSystem;
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _limits = limits ?? throw new ArgumentNullException(nameof(limits));
            _recordBudget = recordBudget ?? throw new ArgumentNullException(nameof(recordBudget));
            _cellBudget = cellBudget ?? throw new ArgumentNullException(nameof(cellBudget));
            _logicalRowBudget = logicalRowBudget ?? throw new ArgumentNullException(nameof(logicalRowBudget));
            _cancellationToken = cancellationToken;
            _activeReadCancellationToken = cancellationToken;
            _sharedStringOwner = sharedStrings as XlsbSharedStringTable;
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            // Records are validated as the reader reaches them; this cursor only carries the
            // position, and budgets are charged by the validating scan.
            _records = new XlsbRecordSliceReader(
                worksheetPart.Buffer,
                limits.MaxRecordBytes,
                recordBudget,
                worksheetPart.DataLength,
                consumeRecordBudget: false);
            _worksheetPart = worksheetPart;
            try {
                AcquireSharedStringReference();
                cancellationToken.ThrowIfCancellationRequested();

                // The schema is fixed before the first row: the declared dimension, widened by
                // the header row, or without one by the first populated row.
                LocatePendingRow();
                Dictionary<int, string?>? headerValues = null;
                int schemaFirstColumn = _dimensionFirstColumn;
                int schemaLastColumn = _dimensionLastColumn;
                if (hasHeaderRow && _hasPendingRow) {
                    _logicalRowBudget.Consume();
                    int headerRowIndex = _pendingRowIndex;
                    headerValues = ReadHeaderRowCells();
                    foreach (int column in headerValues.Keys) {
                        schemaFirstColumn = Math.Min(schemaFirstColumn, column);
                        schemaLastColumn = Math.Max(schemaLastColumn, column);
                    }

                    LocatePendingRow();
                    _nextLogicalRowIndex = checked(headerRowIndex + 1);
                } else if (_hasPendingRow) {
                    PeekPendingRowColumns(out int rowFirstColumn, out int rowLastColumn);
                    schemaFirstColumn = Math.Min(schemaFirstColumn, rowFirstColumn);
                    schemaLastColumn = Math.Max(schemaLastColumn, rowLastColumn);
                    _nextLogicalRowIndex = _pendingRowIndex;
                }

                _hasRows = _hasPendingRow;
                int firstColumn = schemaFirstColumn;
                int lastColumn = schemaLastColumn;
                _firstColumn = firstColumn;
                int fieldCount = lastColumn >= firstColumn
                    ? checked(lastColumn - firstColumn + 1)
                    : 0;
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
                InitializeUtf8TextState();
                _columnTypes = CreateObjectColumnTypes(fieldCount);
                _schemaRows = _options.InferSchema
                    ? BufferSchemaRows()
                    : null;
            } catch {
                worksheetPart.Dispose();
                ReleaseSharedStringReference();
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
            ResetUtf8TextState();
            if (!_pendingRowLocated) {
                LocatePendingRow();
            }
            if (!_hasPendingRow) {
                return false;
            }

            // Typed span clears avoid the general Array.Clear dispatch on every row.
            _kinds.AsSpan().Clear();
            _strings.AsSpan().Clear();
            if (_options.CellValueConverter != null) {
                _customValues.AsSpan().Clear();
            }
            _logicalRowBudget.Consume();
            if (_nextLogicalRowIndex < _pendingRowIndex) {
                // A missing or empty row between populated rows.
                _nextLogicalRowIndex++;
                _hasCurrentRow = true;
                return true;
            }

            int currentRowIndex = _pendingRowIndex;
            ReadPendingRowCells();
            _pendingRowLocated = false;
            _nextLogicalRowIndex = checked(currentRowIndex + 1);
            _hasCurrentRow = true;
            return true;
        }

        public override bool NextResult() => false;

        public override void Close() {
            if (_closed) {
                return;
            }

            _closed = true;
            ReleaseUtf8TextState();
            _worksheetPart.Dispose();
            ReleaseSharedStringReference();
            _schemaRows?.Clear();
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Close();
            }

            base.Dispose(disposing);
        }

        /// <summary>Stores a validated cell through the configured cell-value converter.</summary>
        private void StoreCell(XlsbRecordSlice record, int ordinal) {
            DecodedCell cell = DecodeCell(record);

            _kinds[ordinal] = cell.Kind;
            switch (cell.Kind) {
                case XlsbTabularValueKind.Text:
                case XlsbTabularValueKind.Error:
                    _strings[ordinal] = cell.Text;
                    break;
                case XlsbTabularValueKind.Number:
                case XlsbTabularValueKind.Date:
                case XlsbTabularValueKind.Time:
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
            var cursor = record.CreateCursor();
            int column = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;

            ExcelSerialDateStyle dateStyle = _options.TreatDatesUsingNumberFormat && styleIndex < _dateStyles.Length
                ? _dateStyles[styleIndex] : ExcelSerialDateStyle.None;
            DecodedCell cell;
            switch (record.Type) {
                case BrtCellBlank:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Empty);
                    break;
                case BrtCellRk:
                    cell = NumericCell(column, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), dateStyle);
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
                    cell = NumericCell(column, cursor.ReadDouble(), dateStyle);
                    break;
                case BrtCellSt:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Text) {
                        Text = cursor.ReadWideString(_limits.MaxStringCharacters)
                    };
                    break;
                case BrtCellIsst: {
                    uint sharedStringIndex = cursor.ReadUInt32();
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
                    break;
                case BrtFmlaNum:
                    cell = NumericCell(column, cursor.ReadDouble(), dateStyle);
                    break;
                case BrtFmlaBool:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Boolean) {
                        Boolean = cursor.ReadByte() != 0
                    };
                    break;
                case BrtFmlaError:
                    cell = new DecodedCell(column, XlsbTabularValueKind.Error) {
                        Text = BiffErrorValue.ToText(cursor.ReadByte())
                    };
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }

            return ApplyCellValueConverter(cell, styleIndex, record.Type);
        }

        private static XlsbTabularValueKind NumericKind(ExcelSerialDateStyle style) => style switch {
            ExcelSerialDateStyle.Calendar => XlsbTabularValueKind.Date,
            ExcelSerialDateStyle.Time => XlsbTabularValueKind.Time,
            _ => XlsbTabularValueKind.Number
        };

        private static DecodedCell NumericCell(int column, double number, ExcelSerialDateStyle dateStyle) =>
            new(column, NumericKind(dateStyle)) {
                Number = number
            };

        private string? CellToHeaderText(DecodedCell cell) =>
            cell.Kind switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => cell.Text,
                XlsbTabularValueKind.Number => cell.Number.ToString("R", _options.Culture),
                XlsbTabularValueKind.Boolean => cell.Boolean.ToString(),
                XlsbTabularValueKind.Date or XlsbTabularValueKind.Time => ConvertDate(cell.Number, cell.Kind == XlsbTabularValueKind.Date).ToString(_options.Culture),
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
                XlsbTabularValueKind.Number or XlsbTabularValueKind.Date or XlsbTabularValueKind.Time => CellValues.Number,
                XlsbTabularValueKind.Boolean => CellValues.Boolean,
                XlsbTabularValueKind.Error => CellValues.Error,
                XlsbTabularValueKind.Text when recordType == BrtCellIsst => CellValues.SharedString,
                XlsbTabularValueKind.Text => CellValues.String,
                _ => null
            };
            string? rawText = cell.RawText ?? cell.Kind switch {
                XlsbTabularValueKind.Number or XlsbTabularValueKind.Date or XlsbTabularValueKind.Time =>
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
                    typeHint.ToOfficeEnum(),
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

        private DateTime ConvertDate(double serial, bool calendarStyle) {
            if (LegacyXlsDateSerialConverter.TryConvert(serial, _uses1904DateSystem, out DateTime value, calendarStyle)) {
                return value;
            }

            throw new InvalidCastException($"The XLSB numeric value '{serial}' is not a valid Excel date.");
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

        private static void EnsureFormulaModeSupported(
            int recordType,
            bool useCachedFormulaResult) {
            if (!useCachedFormulaResult && IsFormulaRecord(recordType)) {
                throw new NotSupportedException(
                    "XLSB formula-token projection is not supported when cached formula results are disabled.");
            }
        }

        private void CheckCancellation() {
            if (!CanCancelCurrentRead) {
                return;
            }

            _recordsSinceCancellationCheck++;
            if ((_recordsSinceCancellationCheck & 1023) == 0) {
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

            ref byte payload = ref record.Bytes[record.PayloadOffset];
            uint rowIndex = Unsafe.ReadUnaligned<uint>(ref payload);
            byte extraFlags = Unsafe.Add(ref payload, 10);
            byte flags = Unsafe.Add(ref payload, 11);
            byte phoneticFlags = Unsafe.Add(ref payload, 12);
            uint declaredSpanCount = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref payload, 13));
            if (rowIndex >= ExcelA1.MaxRows
                || (extraFlags & 0xFC) != 0
                || (flags & 0x80) != 0
                || (phoneticFlags & 0xFE) != 0
                || declaredSpanCount > 16) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} contains invalid row metadata.");
            }
            if (record.Size != 17 + checked((int)declaredSpanCount * 8)) {
                throw new InvalidDataException(
                    $"The BrtRowHdr record at offset {record.RecordOffset} has an invalid column-span payload.");
            }

            int previousLast = -1;
            spanCount = checked((int)declaredSpanCount);
            for (int index = 0; index < spanCount; index++) {
                int spanOffset = 17 + index * 8;
                uint firstColumn = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref payload, spanOffset));
                uint lastColumn = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref payload, spanOffset + 4));
                if (firstColumn > lastColumn
                    || lastColumn >= ExcelA1.MaxColumns
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

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ValidateOrdinal(int ordinal) {
            if ((uint)ordinal >= (uint)_headers.Length) {
                ThrowOrdinalOutOfRange(ordinal);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ValidateReadableOrdinal(int ordinal) {
            if (_closed || (uint)ordinal >= (uint)_headers.Length || !_hasCurrentRow) {
                ThrowInvalidReadableOrdinal(ordinal);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowInvalidReadableOrdinal(int ordinal) {
            ThrowIfClosed();
            ValidateOrdinal(ordinal);
            throw new InvalidOperationException("Read must be called before accessing values.");
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowOrdinalOutOfRange(int ordinal) =>
            throw new IndexOutOfRangeException(
                $"Column ordinal {ordinal} is outside 0..{FieldCount - 1}.");

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
                XlsbTabularValueKind.Date or XlsbTabularValueKind.Time => typeof(DateTime),
                XlsbTabularValueKind.Custom => IsMissingCustomValue(_customValues[ordinal])
                    ? typeof(object)
                    : _customValues[ordinal]!.GetType(),
                _ => typeof(object)
            };

        private static bool IsMissingCustomValue(object? value) =>
            value == null || ReferenceEquals(value, DBNull.Value);

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
            Custom,
            Time
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
