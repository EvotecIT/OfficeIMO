#nullable enable

using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.Net;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource : IDisposable {
            private const int InitialBufferSize = 64 * 1024;
            private const int MaximumBufferSize = 64 * 1024 * 1024;
            private const int MaximumIndexedCells = 1_000_000;
            private const int StringCacheSize = 256;
            private const int MaximumCachedStringBytes = 256;

            private readonly ExcelSheetReader _owner;
            private readonly ExcelReadOptions _options;
            private readonly int _firstColumn;
            private readonly int _fieldCount;
            private readonly Utf8StringCacheEntry[] _stringCache;
            private byte[]? _buffer;
            private int[]? _rowIndexes;
            private int[]? _valueStarts;
            private int[]? _valueLengths;
            private int[]? _formulaStarts;
            private int[]? _formulaLengths;
            private int[]? _styleIndexes;
            private byte[]? _cellKinds;
            private int _length;
            private int _rowCount;
            private int _rowCursor;
            private int _currentRowOffset = -1;
            private int _lastDateStyleIndex = -1;
            private bool _lastDateStyleResult;
            private bool _parseFailed;
            private bool _disposed;

            private ExcelUtf8RangeRowSource(
                ExcelSheetReader owner,
                byte[] buffer,
                int length,
                int firstRow,
                int lastRow,
                int firstColumn,
                int fieldCount) {
                _owner = owner;
                _options = owner._opt;
                _buffer = buffer;
                _length = length;
                _firstColumn = firstColumn;
                _fieldCount = fieldCount;

                int rowCapacity = Math.Max(16, Math.Min(lastRow - firstRow + 1, 4096));
                _rowIndexes = ArrayPool<int>.Shared.Rent(rowCapacity);
                int cellCapacity = checked(rowCapacity * fieldCount);
                _valueStarts = ArrayPool<int>.Shared.Rent(cellCapacity);
                _valueLengths = ArrayPool<int>.Shared.Rent(cellCapacity);
                _formulaStarts = ArrayPool<int>.Shared.Rent(cellCapacity);
                _formulaLengths = ArrayPool<int>.Shared.Rent(cellCapacity);
                _styleIndexes = ArrayPool<int>.Shared.Rent(cellCapacity);
                _cellKinds = ArrayPool<byte>.Shared.Rent(cellCapacity);
                _stringCache = new Utf8StringCacheEntry[StringCacheSize];
            }

            internal static bool TryCreate(
                ExcelSheetReader owner,
                int firstRow,
                int lastRow,
                int firstColumn,
                int fieldCount,
                CancellationToken ct,
                out ExcelUtf8RangeRowSource? source) {
                source = null;
                if (owner._opt.CellValueConverter != null
                    || ((long)(lastRow - firstRow + 1) * fieldCount) > MaximumIndexedCells) {
                    return false;
                }

                byte[]? buffer = null;
                try {
                    using var stream = owner._wsPart.GetStream(FileMode.Open, FileAccess.Read);
                    RewindWorksheetStream(stream);
                    if (!TryReadWorksheetBuffer(stream, ct, out buffer, out int length)) {
                        return false;
                    }

                    var candidate = new ExcelUtf8RangeRowSource(
                        owner,
                        buffer!,
                        length,
                        firstRow,
                        lastRow,
                        firstColumn,
                        fieldCount);
                    buffer = null;
                    if (!candidate.TryIndexRows(firstRow, lastRow)) {
                        candidate.Dispose();
                        return false;
                    }

                    source = candidate;
                    return true;
                } finally {
                    if (buffer != null) {
                        ArrayPool<byte>.Shared.Return(buffer);
                    }
                }
            }

            internal bool SelectRow(int rowIndex) {
                EnsureNotDisposed();
                while (_rowCursor < _rowCount && _rowIndexes![_rowCursor] < rowIndex) {
                    _rowCursor++;
                }

                if (_rowCursor >= _rowCount || _rowIndexes![_rowCursor] != rowIndex) {
                    return false;
                }

                _currentRowOffset = checked(_rowCursor * _fieldCount);
                _rowCursor++;
                return true;
            }

            internal void ReadValue(
                int ordinal,
                XmlDataReaderTargetKind targetKind,
                out XmlDataReaderPrimitiveKind primitiveKind,
                out double doubleValue,
                out DateTime dateTimeValue,
                out bool booleanValue,
                out object? objectValue) {
                EnsureNotDisposed();
                primitiveKind = XmlDataReaderPrimitiveKind.None;
                doubleValue = 0;
                dateTimeValue = default;
                booleanValue = false;
                objectValue = null;

                int cellIndex = _currentRowOffset + ordinal;
                if ((Utf8CellKind)_cellKinds![cellIndex] == Utf8CellKind.Missing) {
                    return;
                }

                bool useFormula = _formulaLengths![cellIndex] >= 0
                    && (!_options.UseCachedFormulaResult || _valueLengths![cellIndex] < 0);
                int start = useFormula ? _formulaStarts![cellIndex] : _valueStarts![cellIndex];
                int length = useFormula ? _formulaLengths![cellIndex] : _valueLengths![cellIndex];
                if (length < 0) {
                    return;
                }

                if (useFormula) {
                    objectValue = DecodeString(start, length);
                    return;
                }

                ReadOnlySpan<byte> value = _buffer!.AsSpan(start, length);
                switch ((Utf8CellKind)_cellKinds![cellIndex]) {
                    case Utf8CellKind.SharedString:
                        if (TryParseInt32(value, out int sharedStringIndex)) {
                            objectValue = _owner.GetSharedString(sharedStringIndex);
                        } else {
                            string sharedStringText = DecodeString(start, length);
                            objectValue = TryParseSharedStringIndex(sharedStringText, out sharedStringIndex)
                                ? _owner.GetSharedString(sharedStringIndex)
                                : sharedStringText;
                        }

                        return;
                    case Utf8CellKind.Boolean:
                        bool parsedBoolean = value.Length == 1
                            ? value[0] == (byte)'1'
                            : value.IndexOf((byte)'&') >= 0
                                && string.Equals(DecodeString(start, length), "1", StringComparison.Ordinal);
                        if (targetKind == XmlDataReaderTargetKind.Boolean) {
                            primitiveKind = XmlDataReaderPrimitiveKind.Boolean;
                            booleanValue = parsedBoolean;
                        } else {
                            objectValue = BoxBoolean(parsedBoolean);
                        }
                        return;
                    case Utf8CellKind.Date:
                        string dateText = DecodeString(start, length);
                        objectValue = DateTime.TryParse(dateText, _options.Culture, DateTimeStyles.AssumeLocal, out DateTime parsedDate)
                            ? parsedDate
                            : dateText;
                        return;
                    case Utf8CellKind.String:
                    case Utf8CellKind.Error:
                        objectValue = DecodeString(start, length);
                        return;
                    case Utf8CellKind.Number:
                        ReadNumberValue(cellIndex, value, targetKind, out primitiveKind, out doubleValue, out dateTimeValue, out objectValue);
                        return;
                    default:
                        return;
                }
            }

            public void Dispose() {
                if (_disposed) {
                    return;
                }

                _disposed = true;
                if (_buffer != null) {
                    Array.Clear(_buffer, 0, _length);
                    ArrayPool<byte>.Shared.Return(_buffer);
                    _buffer = null;
                }

                ReturnRowArray(ref _rowIndexes);
                ReturnRowArray(ref _valueStarts);
                ReturnRowArray(ref _valueLengths);
                ReturnRowArray(ref _formulaStarts);
                ReturnRowArray(ref _formulaLengths);
                ReturnRowArray(ref _styleIndexes);
                if (_cellKinds != null) {
                    ArrayPool<byte>.Shared.Return(_cellKinds);
                    _cellKinds = null;
                }
            }

            private static bool TryReadWorksheetBuffer(Stream stream, CancellationToken ct, out byte[]? buffer, out int length) {
                buffer = ArrayPool<byte>.Shared.Rent(InitialBufferSize);
                length = 0;
                while (true) {
                    if (ct.CanBeCanceled) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (length == buffer.Length) {
                        if (buffer.Length >= MaximumBufferSize) {
                            if (stream.ReadByte() >= 0) {
                                ArrayPool<byte>.Shared.Return(buffer);
                                buffer = null;
                                length = 0;
                                return false;
                            }

                            return true;
                        }

                        int nextSize = Math.Min(MaximumBufferSize, checked(buffer.Length * 2));
                        byte[] next = ArrayPool<byte>.Shared.Rent(nextSize);
                        Buffer.BlockCopy(buffer, 0, next, 0, length);
                        ArrayPool<byte>.Shared.Return(buffer);
                        buffer = next;
                    }

                    int read = stream.Read(buffer, length, buffer.Length - length);
                    if (read == 0) {
                        return true;
                    }

                    length += read;
                }
            }

            private bool TryIndexRows(int firstRow, int lastRow) {
                if (!HasSupportedUtf8Encoding()) {
                    return false;
                }

                int position = 0;
                Utf8Tag sheetData = default;
                bool foundSheetData = false;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (!tag.IsEnd && LocalNameEquals(tag, "sheetData")) {
                        sheetData = tag;
                        foundSheetData = true;
                        break;
                    }
                }

                if (!foundSheetData || _parseFailed) {
                    return false;
                }

                if (sheetData.IsEmpty) {
                    return true;
                }

                int nextImplicitRow = 1;
                int previousRow = 0;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (tag.IsEnd && LocalNameEquals(tag, "sheetData")) {
                        return !_parseFailed;
                    }

                    if (tag.IsEnd || !LocalNameEquals(tag, "row")) {
                        return false;
                    }

                    if (!TryGetAttribute(tag, "r", out bool hasRowReference, out int rowReferenceStart, out int rowReferenceLength)) {
                        return false;
                    }

                    int rowIndex = hasRowReference
                        ? ParsePositiveInt(_buffer!, rowReferenceStart, rowReferenceLength)
                        : nextImplicitRow;
                    if (rowIndex <= 0 || rowIndex <= previousRow) {
                        return false;
                    }

                    previousRow = rowIndex;
                    nextImplicitRow = rowIndex + 1;
                    bool includeRow = rowIndex >= firstRow && rowIndex <= lastRow;
                    int rowOffset = -1;
                    if (includeRow) {
                        EnsureRowCapacity(_rowCount + 1);
                        rowOffset = checked(_rowCount * _fieldCount);
                        InitializeMetadataRow(rowOffset);
                    }

                    if (!tag.IsEmpty && !TryIndexRow(ref position, rowOffset)) {
                        return false;
                    }

                    if (includeRow) {
                        _rowIndexes![_rowCount] = rowIndex;
                        _rowCount++;
                    }
                }

                return false;
            }

            private bool TryIndexRow(ref int position, int rowOffset) {
                int nextColumn = 1;
                int previousColumn = 0;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (tag.IsEnd && LocalNameEquals(tag, "row")) {
                        return true;
                    }

                    if (tag.IsEnd || !LocalNameEquals(tag, "c")) {
                        return false;
                    }

                    if (!TryGetCellAttributes(tag, ref nextColumn, out int columnIndex, out Utf8CellKind kind, out int styleIndex)
                        || columnIndex <= previousColumn) {
                        return false;
                    }

                    previousColumn = columnIndex;
                    int ordinal = columnIndex - _firstColumn;
                    int cellIndex = rowOffset >= 0 && (uint)ordinal < (uint)_fieldCount
                        ? rowOffset + ordinal
                        : -1;
                    if (cellIndex >= 0) {
                        _cellKinds![cellIndex] = (byte)kind;
                        _styleIndexes![cellIndex] = styleIndex;
                    }

                    if (!tag.IsEmpty && !TryIndexCell(ref position, cellIndex)) {
                        return false;
                    }
                }

                return false;
            }

            private bool TryIndexCell(ref int position, int cellIndex) {
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (tag.IsEnd && LocalNameEquals(tag, "c")) {
                        return true;
                    }

                    if (tag.IsEnd || (!LocalNameEquals(tag, "v") && !LocalNameEquals(tag, "f"))) {
                        return false;
                    }

                    if (tag.IsEmpty) {
                        if (cellIndex >= 0) {
                            if (LocalNameEquals(tag, "v")) {
                                _valueStarts![cellIndex] = tag.End;
                                _valueLengths![cellIndex] = 0;
                            } else {
                                _formulaStarts![cellIndex] = tag.End;
                                _formulaLengths![cellIndex] = 0;
                            }
                        }
                        continue;
                    }

                    if (!TryReadNextTag(ref position, _length, out Utf8Tag endTag)
                        || !endTag.IsEnd
                        || !LocalNamesEqual(tag, endTag)
                        || ContainsByte(tag.End + 1, endTag.Start, (byte)'<')) {
                        return false;
                    }

                    if (cellIndex >= 0) {
                        int contentStart = tag.End + 1;
                        int contentLength = Math.Max(0, endTag.Start - contentStart);
                        if (LocalNameEquals(tag, "v")) {
                            _valueStarts![cellIndex] = contentStart;
                            _valueLengths![cellIndex] = contentLength;
                        } else {
                            _formulaStarts![cellIndex] = contentStart;
                            _formulaLengths![cellIndex] = contentLength;
                        }
                    }
                }

                return false;
            }

            private void ReadNumberValue(
                int ordinal,
                ReadOnlySpan<byte> value,
                XmlDataReaderTargetKind targetKind,
                out XmlDataReaderPrimitiveKind primitiveKind,
                out double doubleValue,
                out DateTime dateTimeValue,
                out object? objectValue) {
                primitiveKind = XmlDataReaderPrimitiveKind.None;
                doubleValue = 0;
                dateTimeValue = default;
                objectValue = null;
                ReadOnlySpan<byte> trimmed = TrimAsciiWhitespace(value);
                bool dateStyle = IsDateStyle(_styleIndexes![ordinal]);
                if (TryParseDouble(trimmed, out double number)) {
                    if (dateStyle) {
                        DateTime date = _owner.FromExcelSerialDate(number);
                        if (targetKind == XmlDataReaderTargetKind.DateTime) {
                            primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                            dateTimeValue = date;
                        } else {
                            objectValue = date;
                        }

                        return;
                    }

                    if (!_options.NumericAsDecimal) {
                        if (targetKind == XmlDataReaderTargetKind.Int32 || targetKind == XmlDataReaderTargetKind.Double) {
                            primitiveKind = XmlDataReaderPrimitiveKind.Double;
                            doubleValue = number;
                        } else {
                            objectValue = number;
                        }

                        return;
                    }
                }

                if (_options.NumericAsDecimal
                    && Utf8Parser.TryParse(trimmed, out decimal decimalNumber, out int decimalConsumed)
                    && decimalConsumed == trimmed.Length) {
                    objectValue = decimalNumber;
                    return;
                }

                if (trimmed.IndexOf((byte)'&') >= 0) {
                    ReadDecodedNumberValue(
                        dateStyle,
                        DecodeString(_valueStarts![ordinal], _valueLengths![ordinal]),
                        targetKind,
                        out primitiveKind,
                        out doubleValue,
                        out dateTimeValue,
                        out objectValue);
                    return;
                }

                objectValue = DecodeString(_valueStarts![ordinal], _valueLengths![ordinal]);
            }

            private void ReadDecodedNumberValue(
                bool dateStyle,
                string value,
                XmlDataReaderTargetKind targetKind,
                out XmlDataReaderPrimitiveKind primitiveKind,
                out double doubleValue,
                out DateTime dateTimeValue,
                out object? objectValue) {
                primitiveKind = XmlDataReaderPrimitiveKind.None;
                doubleValue = 0;
                dateTimeValue = default;
                objectValue = null;

                bool parsedNumber = TryParseInvariantDoubleFast(value, out double number)
                    || double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);
                if (parsedNumber && dateStyle) {
                    DateTime date = _owner.FromExcelSerialDate(number);
                    if (targetKind == XmlDataReaderTargetKind.DateTime) {
                        primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                        dateTimeValue = date;
                    } else {
                        objectValue = date;
                    }

                    return;
                }

                if (!_options.NumericAsDecimal && parsedNumber) {
                    if (targetKind == XmlDataReaderTargetKind.Int32 || targetKind == XmlDataReaderTargetKind.Double) {
                        primitiveKind = XmlDataReaderPrimitiveKind.Double;
                        doubleValue = number;
                    } else {
                        objectValue = number;
                    }

                    return;
                }

                if (_options.NumericAsDecimal && TryParseRawDecimal(value, _options.Culture, out decimal decimalNumber)) {
                    objectValue = decimalNumber;
                    return;
                }

                objectValue = value;
            }

            private bool IsDateStyle(int styleIndex) {
                if (styleIndex < 0 || !_options.TreatDatesUsingNumberFormat) {
                    return false;
                }

                if (styleIndex == _lastDateStyleIndex) {
                    return _lastDateStyleResult;
                }

                _lastDateStyleIndex = styleIndex;
                _lastDateStyleResult = _owner.Styles.HasDateStyles && _owner.Styles.IsDateLike((uint)styleIndex);
                return _lastDateStyleResult;
            }

            private string DecodeString(int start, int length) {
                if (length <= 0) {
                    return string.Empty;
                }

                if (length <= MaximumCachedStringBytes) {
                    int hash = ComputeHash(_buffer!, start, length);
                    int slot = hash & (StringCacheSize - 1);
                    for (int probe = 0; probe < 8; probe++) {
                        ref Utf8StringCacheEntry entry = ref _stringCache[(slot + probe) & (StringCacheSize - 1)];
                        if (entry.Value == null) {
                            string decoded = DecodeXmlText(start, length);
                            entry = new Utf8StringCacheEntry(hash, start, length, decoded);
                            return decoded;
                        }

                        if (entry.Hash == hash
                            && entry.Length == length
                            && _buffer!.AsSpan(entry.Start, entry.Length).SequenceEqual(_buffer.AsSpan(start, length))) {
                            return entry.Value;
                        }
                    }
                }

                return DecodeXmlText(start, length);
            }

            private string DecodeXmlText(int start, int length) {
                string text = Encoding.UTF8.GetString(_buffer!, start, length);
                if (text.IndexOf('&') >= 0) {
                    text = WebUtility.HtmlDecode(text);
                }

                if (text.IndexOf('\r') >= 0) {
                    text = text.Replace("\r\n", "\n").Replace('\r', '\n');
                }

                return text;
            }

            private void EnsureRowCapacity(int required) {
                int currentCapacity = Math.Min(_rowIndexes!.Length, _valueStarts!.Length / _fieldCount);
                if (required <= currentCapacity) {
                    return;
                }

                int nextCapacity = Math.Min(A1.MaxRows, checked(currentCapacity * 2));
                GrowRowArray(ref _rowIndexes, nextCapacity, _rowCount);
                int currentCellCount = checked(_rowCount * _fieldCount);
                int nextCellCapacity = checked(nextCapacity * _fieldCount);
                GrowRowArray(ref _valueStarts, nextCellCapacity, currentCellCount);
                GrowRowArray(ref _valueLengths, nextCellCapacity, currentCellCount);
                GrowRowArray(ref _formulaStarts, nextCellCapacity, currentCellCount);
                GrowRowArray(ref _formulaLengths, nextCellCapacity, currentCellCount);
                GrowRowArray(ref _styleIndexes, nextCellCapacity, currentCellCount);
                GrowCellKindArray(nextCellCapacity, currentCellCount);
            }

            private void InitializeMetadataRow(int rowOffset) {
                int end = rowOffset + _fieldCount;
                for (int i = rowOffset; i < end; i++) {
                    _cellKinds![i] = (byte)Utf8CellKind.Missing;
                    _valueStarts![i] = 0;
                    _valueLengths![i] = -1;
                    _formulaStarts![i] = 0;
                    _formulaLengths![i] = -1;
                    _styleIndexes![i] = -1;
                }
            }

            private void GrowCellKindArray(int capacity, int count) {
                byte[] next = ArrayPool<byte>.Shared.Rent(capacity);
                Array.Copy(_cellKinds!, next, count);
                ArrayPool<byte>.Shared.Return(_cellKinds!);
                _cellKinds = next;
            }

            private static void GrowRowArray(ref int[]? values, int capacity, int count) {
                int[] next = ArrayPool<int>.Shared.Rent(capacity);
                Array.Copy(values!, next, count);
                ArrayPool<int>.Shared.Return(values!);
                values = next;
            }

            private static void ReturnRowArray(ref int[]? values) {
                if (values != null) {
                    ArrayPool<int>.Shared.Return(values);
                    values = null;
                }
            }

            private static bool TryParseDouble(ReadOnlySpan<byte> value, out double result) {
                return Utf8Parser.TryParse(value, out result, out int consumed) && consumed == value.Length;
            }

            private static bool TryParseInt32(ReadOnlySpan<byte> value, out int result) {
                ReadOnlySpan<byte> trimmed = TrimAsciiWhitespace(value);
                return Utf8Parser.TryParse(trimmed, out result, out int consumed) && consumed == trimmed.Length;
            }

            private static bool ParseBoolean(ReadOnlySpan<byte> value) {
                return value.Length == 1 && value[0] == (byte)'1';
            }

            private static ReadOnlySpan<byte> TrimAsciiWhitespace(ReadOnlySpan<byte> value) {
                int start = 0;
                int end = value.Length;
                while (start < end && IsAsciiWhitespace(value[start])) start++;
                while (end > start && IsAsciiWhitespace(value[end - 1])) end--;
                return value.Slice(start, end - start);
            }

            private static int ParsePositiveInt(byte[] data, int start, int length) {
                if (length <= 0) return 0;
                int value = 0;
                for (int i = 0; i < length; i++) {
                    int digit = data[start + i] - (byte)'0';
                    if ((uint)digit > 9U || value > (int.MaxValue - digit) / 10) return 0;
                    value = (value * 10) + digit;
                }

                return value;
            }

            private static bool TryParseNonNegativeInt(byte[] data, int start, int length, out int value) {
                value = 0;
                if (length <= 0) return false;
                for (int i = 0; i < length; i++) {
                    int digit = data[start + i] - (byte)'0';
                    if ((uint)digit > 9U || value > (int.MaxValue - digit) / 10) {
                        value = 0;
                        return false;
                    }

                    value = (value * 10) + digit;
                }

                return true;
            }

            private static int ParseColumnIndex(byte[] data, int start, int length) {
                int column = 0;
                int position = 0;
                while (position < length) {
                    byte current = data[start + position];
                    int letter = current >= (byte)'a' && current <= (byte)'z'
                        ? current - (byte)'a' + 1
                        : current >= (byte)'A' && current <= (byte)'Z'
                            ? current - (byte)'A' + 1
                            : 0;
                    if (letter == 0) break;
                    if (column > (int.MaxValue - letter) / 26) return 0;
                    column = (column * 26) + letter;
                    position++;
                }

                if (column <= 0 || position >= length) return 0;
                for (; position < length; position++) {
                    byte current = data[start + position];
                    if (current < (byte)'0' || current > (byte)'9') return 0;
                }

                return column;
            }

            private static int ComputeHash(byte[] data, int start, int length) {
                unchecked {
                    uint hash = 2166136261;
                    for (int i = 0; i < length; i++) {
                        hash = (hash ^ data[start + i]) * 16777619;
                    }

                    return (int)hash;
                }
            }

            private int IndexOfAsciiIgnoreCase(int start, int length, string value) {
                int end = start + length - value.Length;
                for (int i = start; i <= end; i++) {
                    if (AsciiEqualsIgnoreCase(i, value.Length, value)) return i;
                }

                return -1;
            }

            private bool AsciiEquals(int start, int length, string value) {
                if (length != value.Length) return false;
                for (int i = 0; i < length; i++) {
                    if (_buffer![start + i] != (byte)value[i]) return false;
                }

                return true;
            }

            private bool AsciiEqualsIgnoreCase(int start, int length, string value) {
                if (length != value.Length) return false;
                for (int i = 0; i < length; i++) {
                    byte current = _buffer![start + i];
                    if (current >= (byte)'A' && current <= (byte)'Z') current = (byte)(current + 32);
                    char expected = value[i];
                    if (expected >= 'A' && expected <= 'Z') expected = (char)(expected + 32);
                    if (current != (byte)expected) return false;
                }

                return true;
            }

            private bool LocalNameEquals(Utf8Tag tag, string name) =>
                AsciiEquals(tag.LocalNameStart, tag.NameEnd - tag.LocalNameStart, name);

            private bool LocalNamesEqual(Utf8Tag first, Utf8Tag second) {
                int firstLength = first.NameEnd - first.LocalNameStart;
                int secondLength = second.NameEnd - second.LocalNameStart;
                return firstLength == secondLength
                    && _buffer!.AsSpan(first.LocalNameStart, firstLength).SequenceEqual(_buffer.AsSpan(second.LocalNameStart, secondLength));
            }

            private bool ContainsByte(int start, int end, byte value) =>
                end > start && _buffer!.AsSpan(start, end - start).IndexOf(value) >= 0;

            private int IndexOfSequence(int start, int limit, params byte[] sequence) {
                int end = limit - sequence.Length;
                for (int i = start; i <= end; i++) {
                    bool matches = true;
                    for (int j = 0; j < sequence.Length; j++) {
                        if (_buffer![i + j] != sequence[j]) {
                            matches = false;
                            break;
                        }
                    }

                    if (matches) return i;
                }

                return -1;
            }

            private static bool IsAsciiWhitespace(byte value) =>
                value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n';

            private static bool IsTagNameTerminator(byte value) =>
                IsAsciiWhitespace(value) || value == (byte)'/' || value == (byte)'>';

            private static bool IsAttributeNameTerminator(byte value) =>
                IsAsciiWhitespace(value) || value == (byte)'=' || value == (byte)'/' || value == (byte)'>';

            private void EnsureNotDisposed() {
                if (_disposed) {
                    throw new ObjectDisposedException(nameof(ExcelUtf8RangeRowSource));
                }
            }

            private enum Utf8CellKind : byte {
                Missing,
                Number,
                SharedString,
                String,
                Boolean,
                Date,
                Error
            }

            private readonly struct Utf8Tag {
                internal Utf8Tag(int start, int end, int nameStart, int nameEnd, int localNameStart, bool isEnd, bool isEmpty) {
                    Start = start;
                    End = end;
                    NameStart = nameStart;
                    NameEnd = nameEnd;
                    LocalNameStart = localNameStart;
                    IsEnd = isEnd;
                    IsEmpty = isEmpty;
                }

                internal int Start { get; }
                internal int End { get; }
                internal int NameStart { get; }
                internal int NameEnd { get; }
                internal int LocalNameStart { get; }
                internal bool IsEnd { get; }
                internal bool IsEmpty { get; }
            }

            private readonly struct Utf8StringCacheEntry {
                internal Utf8StringCacheEntry(int hash, int start, int length, string value) {
                    Hash = hash;
                    Start = start;
                    Length = length;
                    Value = value;
                }

                internal int Hash { get; }
                internal int Start { get; }
                internal int Length { get; }
                internal string? Value { get; }
            }
        }
    }
}
