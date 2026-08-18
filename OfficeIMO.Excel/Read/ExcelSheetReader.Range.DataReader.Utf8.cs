#nullable enable

using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.Net;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource : IDisposable {
            private const int InitialBufferSize = 64 * 1024;
            private const int MaximumBufferSize = 64 * 1024 * 1024;
            private const int MaximumIndexedCells = 1_000_000;
            private const int StringCacheSize = 256;
            private const int MaximumCachedStringBytes = 256;
            private const byte DateStyleCellKindFlag = 0x80;
            private const byte CellKindMask = 0x7F;
            private const int SharedStringIndexValueLength = -2;

            private readonly ExcelSheetReader _owner;
            private readonly ExcelReadOptions _options;
            private int _firstColumn;
            private int _fieldCount;
            private readonly Utf8StringCacheEntry[] _stringCache;
            private byte[]? _buffer;
            private int[]? _rowIndexes;
            private int[]? _valueStarts;
            private int[]? _valueLengths;
            private int[]? _formulaStarts;
            private int[]? _formulaLengths;
            private byte[]? _cellKinds;
            private int _length;
            private int _rowCount;
            private int _rowCursor;
            private int _currentRowOffset = -1;
            private int _minimumCellRow = int.MaxValue;
            private int _maximumCellRow;
            private int _minimumCellColumn = int.MaxValue;
            private int _maximumCellColumn;
            private int _sheetDataContentStart = -1;
            private int _sheetDataEndTagStart = -1;
            private int _repeatedRowSuffixStart = -1;
            private int _repeatedRowSuffixLength;
            private bool _repeatedRowIsEmpty;
            private bool _sheetDataSupportsFastValidation = true;
            private int _lastDateStyleIndex = -1;
            private bool _lastDateStyleResult;
            private bool _cellsFitWithinRange = true;
            private bool _parseFailed;
            private bool _disposed;

            private ExcelUtf8RangeRowSource(
                ExcelSheetReader owner,
                byte[] buffer,
                int length) {
                _owner = owner;
                _options = owner._opt;
                _buffer = buffer;
                _length = length;
                _stringCache = new Utf8StringCacheEntry[StringCacheSize];
            }

            private ExcelUtf8RangeRowSource(
                ExcelSheetReader owner,
                byte[] buffer,
                int length,
                int firstRow,
                int lastRow,
                int firstColumn,
                int fieldCount)
                : this(owner, buffer, length) {
                InitializeRange(firstRow, lastRow, firstColumn, fieldCount);
            }

            private void InitializeRange(
                int firstRow,
                int lastRow,
                int firstColumn,
                int fieldCount) {
                _firstColumn = firstColumn;
                _fieldCount = fieldCount;

                int rowCapacity = Math.Max(16, Math.Min(lastRow - firstRow + 1, 4096));
                _rowIndexes = ArrayPool<int>.Shared.Rent(rowCapacity);
                int cellCapacity = checked(rowCapacity * fieldCount);
                _valueStarts = ArrayPool<int>.Shared.Rent(cellCapacity);
                _valueLengths = ArrayPool<int>.Shared.Rent(cellCapacity);
                _cellKinds = ArrayPool<byte>.Shared.Rent(cellCapacity);
            }

            internal static bool TryCreateForUsedRange(
                ExcelSheetReader owner,
                CancellationToken ct,
                out ExcelUtf8RangeRowSource? source,
                out int declaredFirstColumn) {
                source = null;
                declaredFirstColumn = 0;
                if (owner._opt.CellValueConverter != null) {
                    return false;
                }

                byte[]? buffer = null;
                try {
                    if (!TryRentWorksheetBuffer(owner, ct, out buffer, out int length)) {
                        return false;
                    }

                    var candidate = new ExcelUtf8RangeRowSource(owner, buffer!, length);
                    buffer = null;
                    try {
                        if (!candidate.TryGetDeclaredRange(
                                ct,
                                out int firstRow,
                                out int firstColumn,
                                out int lastRow,
                                out int lastColumn)) {
                            candidate.Dispose();
                            return false;
                        }

                        int fieldCount = lastColumn - firstColumn + 1;
                        long rowCount = (long)lastRow - firstRow + 1L;
                        long cellCount = rowCount * fieldCount;
                        if (fieldCount <= 0
                            || fieldCount > owner._opt.MaxDataReaderColumns
                            || cellCount > owner._opt.MaxDataReaderBufferedCells
                            || cellCount > MaximumIndexedCells) {
                            candidate.Dispose();
                            return false;
                        }

                        candidate.InitializeRange(firstRow, lastRow, firstColumn, fieldCount);
                        if (!candidate.TryIndexRows(firstRow, lastRow, ct)) {
                            candidate.Dispose();
                            return false;
                        }

                        if (!candidate.IsCanonicalWorksheetXmlFullyValidated(ct)) {
                            candidate.ValidateBufferedWorksheetXml(ct);
                        }

                        ct.ThrowIfCancellationRequested();
                        declaredFirstColumn = firstColumn;
                        source = candidate;
                        return true;
                    } catch {
                        candidate.Dispose();
                        throw;
                    }
                } finally {
                    if (buffer != null) {
                        ArrayPool<byte>.Shared.Return(buffer);
                    }
                }
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
                    if (!TryRentWorksheetBuffer(owner, ct, out buffer, out int length)) {
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
                    try {
                        if (!candidate.TryIndexRows(firstRow, lastRow, ct)) {
                            candidate.Dispose();
                            return false;
                        }

                        if (!candidate.IsCanonicalWorksheetXmlFullyValidated(ct)) {
                            candidate.ValidateBufferedWorksheetXml(ct);
                        }
                    } catch {
                        candidate.Dispose();
                        throw;
                    }

                    ct.ThrowIfCancellationRequested();
                    source = candidate;
                    return true;
                } finally {
                    if (buffer != null) {
                        ArrayPool<byte>.Shared.Return(buffer);
                    }
                }
            }

            private static bool TryRentWorksheetBuffer(
                ExcelSheetReader owner,
                CancellationToken ct,
                out byte[]? buffer,
                out int length) {
                if (owner.TryReadWorksheetPartBuffer(
                        MaximumBufferSize,
                        ct,
                        out buffer,
                        out length)) {
                    return true;
                }

                owner.RequireSdkWorksheetPart();
                using var stream = owner._wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                return TryReadWorksheetBuffer(stream, ct, out buffer, out length);
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

            internal bool CellsFitWithinRange => _cellsFitWithinRange;

            internal bool TryGetUsedBounds(
                out int firstRow,
                out int firstColumn,
                out int lastRow,
                out int lastColumn) {
                firstRow = _minimumCellRow;
                firstColumn = _minimumCellColumn;
                lastRow = _maximumCellRow;
                lastColumn = _maximumCellColumn;
                return lastRow > 0 && lastColumn > 0;
            }

            internal void ReadValue(
                int ordinal,
                XmlDataReaderTargetKind targetKind,
                out XmlDataReaderPrimitiveKind primitiveKind,
                out double doubleValue,
                out DateTime dateTimeValue,
                out bool booleanValue,
                out bool isFormulaText,
                out bool deferObjectMaterialization,
                out object? objectValue) {
                EnsureNotDisposed();
                primitiveKind = XmlDataReaderPrimitiveKind.None;
                doubleValue = 0;
                dateTimeValue = default;
                booleanValue = false;
                isFormulaText = false;
                deferObjectMaterialization = false;
                objectValue = null;

                int cellIndex = _currentRowOffset + ordinal;
                byte encodedCellKind = _cellKinds![cellIndex];
                Utf8CellKind cellKind = (Utf8CellKind)(encodedCellKind & CellKindMask);
                if (cellKind == Utf8CellKind.Missing) {
                    return;
                }

                bool useFormula = _formulaLengths != null
                    && _formulaLengths[cellIndex] >= 0
                    && (!_options.UseCachedFormulaResult
                        || (_valueLengths![cellIndex] < 0
                            && _valueLengths[cellIndex] != SharedStringIndexValueLength));
                int start = useFormula ? _formulaStarts![cellIndex] : _valueStarts![cellIndex];
                int length = useFormula ? _formulaLengths![cellIndex] : _valueLengths![cellIndex];
                if (!useFormula
                    && cellKind == Utf8CellKind.SharedString
                    && length == SharedStringIndexValueLength) {
                    objectValue = _owner.GetSharedString(start);
                    return;
                }
                if (length < 0) {
                    return;
                }

                if (useFormula) {
                    isFormulaText = true;
                    objectValue = DecodeString(start, length);
                    return;
                }

                ReadOnlySpan<byte> value = _buffer!.AsSpan(start, length);
                switch (cellKind) {
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
                        if (targetKind == XmlDataReaderTargetKind.String) {
                            objectValue = parsedBoolean.ToString();
                        } else if (targetKind == XmlDataReaderTargetKind.Boolean) {
                            primitiveKind = XmlDataReaderPrimitiveKind.Boolean;
                            booleanValue = parsedBoolean;
                        } else {
                            objectValue = BoxBoolean(parsedBoolean);
                        }
                        return;
                    case Utf8CellKind.Date:
                        string dateText = DecodeString(start, length);
                        objectValue = targetKind == XmlDataReaderTargetKind.String
                            ? dateText
                            : DateTime.TryParse(dateText, _options.Culture, DateTimeStyles.AssumeLocal, out DateTime parsedDate)
                            ? parsedDate
                            : dateText;
                        return;
                    case Utf8CellKind.String:
                    case Utf8CellKind.InlineString:
                    case Utf8CellKind.Error:
                        objectValue = DecodeString(start, length);
                        return;
                    case Utf8CellKind.Number:
                        bool dateStyle = (encodedCellKind & DateStyleCellKindFlag) != 0;
                        deferObjectMaterialization =
                            targetKind == XmlDataReaderTargetKind.Numeric
                            && dateStyle;
                        ReadNumberValue(cellIndex, value, dateStyle, targetKind, out primitiveKind, out doubleValue, out dateTimeValue, out objectValue);
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

            private bool TryIndexRows(int firstRow, int lastRow, CancellationToken ct) {
                ct.ThrowIfCancellationRequested();
                if (!HasSupportedUtf8Encoding()) {
                    return false;
                }

                Span<int> namespacePrefixStarts = stackalloc int[MaximumFastXmlAttributes];
                Span<int> namespacePrefixLengths = stackalloc int[MaximumFastXmlAttributes];
                Span<int> namespaceUriStarts = stackalloc int[MaximumFastXmlAttributes];
                Span<int> namespaceUriLengths = stackalloc int[MaximumFastXmlAttributes];
                int namespacePrefixCount = 0;
                bool hasCanonicalRootNamespaces = false;
                int position = 0;
                Utf8Tag sheetData = default;
                bool foundSheetData = false;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    ct.ThrowIfCancellationRequested();
                    if (!tag.IsEnd
                        && IsUnprefixedTag(tag)
                        && LocalNameEquals(tag, "worksheet")) {
                        hasCanonicalRootNamespaces = ValidateCanonicalNamespaceUsage(
                            tag,
                            isRootStartTag: true,
                            namespacePrefixStarts,
                            namespacePrefixLengths,
                            namespaceUriStarts,
                            namespaceUriLengths,
                            ref namespacePrefixCount);
                    }
                    if (!tag.IsEnd && IsUnprefixedTag(tag) && LocalNameEquals(tag, "sheetData")) {
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
                _sheetDataContentStart = sheetData.End + 1;

                int nextImplicitRow = 1;
                int previousRow = 0;
                bool repeatedRowShapeValidated = false;
                while (position < _length) {
                    ct.ThrowIfCancellationRequested();
                    int tagSearchStart = position;
                    Utf8Tag tag = default;
                    int rowIndex = 0;
                    bool repeatedRow = repeatedRowShapeValidated
                        && TryReadRepeatedRowStartTag(
                            ref position,
                            out tag,
                            out rowIndex);
                    if (!repeatedRow
                        && !TryReadNextTag(ref position, _length, out tag)) {
                        return false;
                    }
                    if (!repeatedRow && ContainsNonWhitespace(tagSearchStart, tag.Start)) {
                        _sheetDataSupportsFastValidation = false;
                    }
                    if (tag.IsEnd && IsUnprefixedTag(tag) && LocalNameEquals(tag, "sheetData")) {
                        _sheetDataEndTagStart = tag.Start;
                        return !_parseFailed;
                    }

                    if (tag.IsEnd || !IsUnprefixedTag(tag) || !LocalNameEquals(tag, "row")) {
                        return false;
                    }

                    if (!repeatedRow) {
                        bool rowSupportsFastValidation = ValidateCanonicalTagAttributes(
                            tag,
                            out bool rowDeclaresDefaultNamespace,
                            out _,
                            out bool rowHasPrefixedAttributes)
                            && !rowDeclaresDefaultNamespace
                            && (!rowHasPrefixedAttributes
                                || hasCanonicalRootNamespaces
                                && ValidateCanonicalNamespaceUsage(
                                    tag,
                                    isRootStartTag: false,
                                    namespacePrefixStarts,
                                    namespacePrefixLengths,
                                    namespaceUriStarts,
                                    namespaceUriLengths,
                                    ref namespacePrefixCount));
                        if (!rowSupportsFastValidation) {
                            _sheetDataSupportsFastValidation = false;
                        }

                        if (!TryGetAttribute(tag, "r", out bool hasRowReference, out int rowReferenceStart, out int rowReferenceLength)) {
                            return false;
                        }

                        rowIndex = hasRowReference
                            ? ParsePositiveInt(_buffer!, rowReferenceStart, rowReferenceLength)
                            : nextImplicitRow;
                        if (rowSupportsFastValidation) {
                            repeatedRowShapeValidated = TryCaptureRepeatedRowShape(tag, rowIndex);
                        }
                    }
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

                    if (!tag.IsEmpty) {
                        int rowContentStart = position;
                        bool indexedCanonicalRow = rowOffset >= 0
                            && TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
                        if (!indexedCanonicalRow) {
                            position = rowContentStart;
                            if (rowOffset >= 0) {
                                InitializeMetadataRow(rowOffset);
                            }
                            if (!TryIndexRow(ref position, rowIndex, rowOffset, includeRow, ct)) {
                                return false;
                            }
                        }
                    }

                    if (includeRow) {
                        _rowIndexes![_rowCount] = rowIndex;
                        _rowCount++;
                    }

                }

                return false;
            }

            private bool TryIndexRow(
                ref int position,
                int rowIndex,
                int rowOffset,
                bool rowWithinRange,
                CancellationToken ct) {
                int nextColumn = 1;
                int previousColumn = 0;
                while (position < _length) {
                    ct.ThrowIfCancellationRequested();
                    int originalPosition = position;
                    int originalNextColumn = nextColumn;
                    bool compactTag = TryReadCompactCellStartTag(
                        ref position,
                        ref nextColumn,
                        out Utf8Tag tag,
                        out int columnIndex,
                        out Utf8CellKind kind,
                        out int styleIndex);
                    if (!compactTag) {
                        position = originalPosition;
                        nextColumn = originalNextColumn;
                        if (!TryReadNextTag(ref position, _length, out tag)) {
                            return false;
                        }
                        if (tag.IsEnd && IsUnprefixedTag(tag) && LocalNameEquals(tag, "row")) {
                            if (ContainsNonWhitespace(originalPosition, tag.Start)) {
                                _sheetDataSupportsFastValidation = false;
                            }
                            UpdateUsedBounds(rowIndex, previousColumn);
                            return true;
                        }
                        _sheetDataSupportsFastValidation = false;
                        if (tag.IsEnd
                            || !IsUnprefixedTag(tag)
                            || !LocalNameEquals(tag, "c")
                            || !TryGetCellAttributes(tag, ref nextColumn, out columnIndex, out kind, out styleIndex)) {
                            return false;
                        }
                    }
                    if (columnIndex <= previousColumn) {
                        return false;
                    }

                    int firstColumnInRow = previousColumn == 0 ? columnIndex : 0;
                    previousColumn = columnIndex;
                    if (firstColumnInRow > 0
                        && (_minimumCellColumn == int.MaxValue || firstColumnInRow < _minimumCellColumn)) {
                        _minimumCellColumn = firstColumnInRow;
                    }
                    int ordinal = columnIndex - _firstColumn;
                    if (!rowWithinRange || (uint)ordinal >= (uint)_fieldCount) {
                        _cellsFitWithinRange = false;
                    }

                    int cellIndex = rowOffset >= 0 && (uint)ordinal < (uint)_fieldCount
                        ? rowOffset + ordinal
                        : -1;

                    bool sharedFormulaFollower = false;
                    bool hasCachedValue = false;
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!tag.IsEmpty) {
                        int contentPosition = position;
                        if (!TryIndexCompactValueCell(
                                ref position,
                                cellIndex,
                                out hasCachedValue,
                                out valueStart,
                                out valueLength)) {
                            _sheetDataSupportsFastValidation = false;
                            position = contentPosition;
                            if (!TryIndexCell(
                                    ref position,
                                    cellIndex,
                                    kind,
                                    out sharedFormulaFollower,
                                    out hasCachedValue,
                                    out valueStart,
                                    out valueLength)) {
                                return false;
                            }
                        }
                    }

                    int sharedStringIndex = ValidateIndexedCell(
                        rowIndex,
                        columnIndex,
                        kind,
                        styleIndex,
                        sharedFormulaFollower,
                        hasCachedValue,
                        valueStart,
                        valueLength);
                    if (cellIndex >= 0) {
                        if (sharedStringIndex >= 0) {
                            _valueStarts![cellIndex] = sharedStringIndex;
                            _valueLengths![cellIndex] = SharedStringIndexValueLength;
                        }
                        _cellKinds![cellIndex] = EncodeCellKind(kind, styleIndex);
                    }
                }

                return false;
            }

            private void UpdateUsedBounds(int rowIndex, int lastColumnInRow) {
                if (lastColumnInRow <= 0) {
                    return;
                }

                if (_minimumCellRow == int.MaxValue) {
                    _minimumCellRow = rowIndex;
                }
                _maximumCellRow = rowIndex;
                if (lastColumnInRow > _maximumCellColumn) {
                    _maximumCellColumn = lastColumnInRow;
                }
            }

            private bool TryIndexSimpleInlineStringCell(ref int position, Utf8Tag inlineStringTag, int cellIndex) {
                if (inlineStringTag.IsEmpty) {
                    if (cellIndex >= 0) {
                        _valueStarts![cellIndex] = inlineStringTag.End;
                        _valueLengths![cellIndex] = 0;
                    }

                    return TryReadInlineStringCellEnd(ref position, inlineStringTag.End + 1);
                }

                int contentBoundary = inlineStringTag.End + 1;
                if (!TryReadNextTag(ref position, _length, out Utf8Tag textTag)
                    || ContainsNonWhitespace(contentBoundary, textTag.Start)
                    || textTag.IsEnd
                    || !IsUnprefixedTag(textTag)
                    || !LocalNameEquals(textTag, "t")) {
                    return false;
                }

                int valueStart = textTag.End;
                int valueLength = 0;
                int nextBoundary = textTag.End + 1;
                if (!textTag.IsEmpty) {
                    if (!TryReadNextTag(ref position, _length, out Utf8Tag textEndTag)
                        || !textEndTag.IsEnd
                        || !IsUnprefixedTag(textEndTag)
                        || !LocalNamesEqual(textTag, textEndTag)
                        || ContainsByte(textTag.End + 1, textEndTag.Start, (byte)'<')) {
                        return false;
                    }

                    valueStart = textTag.End + 1;
                    valueLength = Math.Max(0, textEndTag.Start - valueStart);
                    nextBoundary = textEndTag.End + 1;
                }

                if (!TryReadNextTag(ref position, _length, out Utf8Tag inlineStringEndTag)
                    || ContainsNonWhitespace(nextBoundary, inlineStringEndTag.Start)
                    || !inlineStringEndTag.IsEnd
                    || !IsUnprefixedTag(inlineStringEndTag)
                    || !LocalNamesEqual(inlineStringTag, inlineStringEndTag)) {
                    return false;
                }

                if (cellIndex >= 0) {
                    _valueStarts![cellIndex] = valueStart;
                    _valueLengths![cellIndex] = valueLength;
                }

                return TryReadInlineStringCellEnd(ref position, inlineStringEndTag.End + 1);
            }

            private bool TryReadInlineStringCellEnd(ref int position, int contentBoundary) {
                return TryReadNextTag(ref position, _length, out Utf8Tag cellEndTag)
                    && !ContainsNonWhitespace(contentBoundary, cellEndTag.Start)
                    && cellEndTag.IsEnd
                    && IsUnprefixedTag(cellEndTag)
                    && LocalNameEquals(cellEndTag, "c");
            }

            private void ReadNumberValue(
                int ordinal,
                ReadOnlySpan<byte> value,
                bool dateStyle,
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
                if (targetKind == XmlDataReaderTargetKind.String) {
                    if (dateStyle && TryParseDouble(trimmed, out double serialDate)) {
                        objectValue = _owner.FromExcelSerialDate(serialDate).ToString(_options.Culture);
                    } else {
                        objectValue = DecodeString(_valueStarts![ordinal], _valueLengths![ordinal]);
                    }

                    return;
                }

                if (TryParseDouble(trimmed, out double number)) {
                    if (dateStyle
                        && targetKind != XmlDataReaderTargetKind.Numeric) {
                        DateTime date = _owner.FromExcelSerialDate(number);
                        if (targetKind == XmlDataReaderTargetKind.DateTime) {
                            primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                            dateTimeValue = date;
                        } else {
                            objectValue = date;
                        }

                        return;
                    }

                    if (targetKind == XmlDataReaderTargetKind.Numeric) {
                        primitiveKind = XmlDataReaderPrimitiveKind.Double;
                        doubleValue = number;
                        return;
                    }

                    if (!_options.NumericAsDecimal) {
                        objectValue = number;
                        return;
                    }

                    if (TryConvertExcelNumberToDecimal(number, out decimal decimalNumber)) {
                        objectValue = decimalNumber;
                        return;
                    }

                    objectValue = number;
                    return;
                }

                ReadDecodedNumberValue(
                    dateStyle,
                    DecodeString(_valueStarts![ordinal], _valueLengths![ordinal]),
                    targetKind,
                    out primitiveKind,
                    out doubleValue,
                    out dateTimeValue,
                    out objectValue);
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

                bool parsedNumber = double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, _options.Culture, out double number)
                    || TryParseInvariantDoubleFast(value, out number)
                    || double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);
                if (parsedNumber
                    && dateStyle
                    && targetKind != XmlDataReaderTargetKind.Numeric) {
                    DateTime date = _owner.FromExcelSerialDate(number);
                    if (targetKind == XmlDataReaderTargetKind.DateTime) {
                        primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                        dateTimeValue = date;
                    } else {
                        objectValue = date;
                    }

                    return;
                }

                if (parsedNumber
                    && targetKind == XmlDataReaderTargetKind.Numeric) {
                    primitiveKind = XmlDataReaderPrimitiveKind.Double;
                    doubleValue = number;
                    return;
                }

                if (!_options.NumericAsDecimal && parsedNumber) {
                    objectValue = number;
                    return;
                }

                if (_options.NumericAsDecimal && parsedNumber) {
                    objectValue = TryConvertExcelNumberToDecimal(number, out decimal decimalNumber)
                        ? decimalNumber
                        : number;
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
                if (_formulaStarts != null) {
                    GrowRowArray(ref _formulaStarts, nextCellCapacity, currentCellCount);
                    GrowRowArray(ref _formulaLengths, nextCellCapacity, currentCellCount);
                }
                GrowCellKindArray(nextCellCapacity, currentCellCount);
            }

            private void InitializeMetadataRow(int rowOffset) {
                int end = rowOffset + _fieldCount;
                for (int i = rowOffset; i < end; i++) {
                    _cellKinds![i] = (byte)Utf8CellKind.Missing;
                    _valueStarts![i] = 0;
                    _valueLengths![i] = -1;
                    if (_formulaStarts != null) {
                        _formulaStarts[i] = 0;
                        _formulaLengths![i] = -1;
                    }
                }
            }

            private byte EncodeCellKind(Utf8CellKind kind, int styleIndex) {
                byte encoded = (byte)kind;
                if (kind == Utf8CellKind.Number && styleIndex >= 0 && IsDateStyle(styleIndex)) {
                    encoded |= DateStyleCellKindFlag;
                }
                return encoded;
            }

            private void EnsureFormulaMetadata() {
                if (_formulaStarts != null) {
                    return;
                }

                int capacity = _valueStarts!.Length;
                _formulaStarts = ArrayPool<int>.Shared.Rent(capacity);
                _formulaLengths = ArrayPool<int>.Shared.Rent(capacity);
                int initializedCellCount = checked((_rowCount + 1) * _fieldCount);
                Array.Clear(_formulaStarts, 0, initializedCellCount);
                for (int index = 0; index < initializedCellCount; index++) {
                    _formulaLengths[index] = -1;
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
                return ExcelUtf8NumberParser.TryParseDouble(value, out result);
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

            private bool ContainsNonWhitespace(int start, int end) {
                for (int i = start; i < end; i++) {
                    if (!IsAsciiWhitespace(_buffer![i])) {
                        return true;
                    }
                }

                return false;
            }

            private int IndexOfSequence(int start, int limit, byte first, byte second) {
                int end = limit - 2;
                for (int i = start; i <= end; i++) {
                    if (_buffer![i] == first && _buffer[i + 1] == second) return i;
                }

                return -1;
            }

            private int IndexOfSequence(int start, int limit, byte first, byte second, byte third) {
                int end = limit - 3;
                for (int i = start; i <= end; i++) {
                    if (_buffer![i] == first && _buffer[i + 1] == second && _buffer[i + 2] == third) return i;
                }

                return -1;
            }

            private static bool IsAsciiWhitespace(byte value) =>
                value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n';

            private static bool IsUnprefixedTag(Utf8Tag tag) =>
                tag.NameStart == tag.LocalNameStart;

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
                InlineString,
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
