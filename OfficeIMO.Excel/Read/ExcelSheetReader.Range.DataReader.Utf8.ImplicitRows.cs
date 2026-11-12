namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Commits the provisional next row only after every cell has been proven
            /// to omit its coordinate. A later explicit cell reference may determine
            /// the row number, so any incomplete attempt restores the discovery state
            /// before the ordinary coordinate lookahead runs.
            /// A committed row initializes every metadata slot; callers reset partial
            /// metadata before indexing an unsuccessful attempt through the fallback.
            /// </summary>
            private bool TryIndexCanonicalImplicitDenseRow(
                ref int position,
                int rowIndex,
                int rowOffset,
                CancellationToken ct) {
                int originalPosition = position;
                if (_cachedImplicitRowReady) {
                    try {
                        if (TryIndexCachedImplicitRow(ref position, rowIndex, rowOffset, ct)) return true;
                    } catch (InvalidDataException) {
                        // A later explicit cell can still change this provisional row.
                        // Ordinary lookahead repeats validation with the actual reference.
                    }
                    // A changing layout keeps the adaptive per-cell parser. Do not
                    // repeatedly scan prefixes through a rejected whole-row shape.
                    _cachedImplicitRowRejected = true;
                    position = originalPosition;
                }
                _cachedImplicitRowReady = false;
                // Cached misses leave bounds untouched. Only the adaptive parser
                // below needs a rollback snapshot; the usual cached row does not.
                int minimumCellRow = _minimumCellRow;
                int maximumCellRow = _maximumCellRow;
                int minimumCellColumn = _minimumCellColumn;
                int maximumCellColumn = _maximumCellColumn;
                try {
                    if (TryIndexCanonicalImplicitDenseRowCore(ref position, rowIndex, rowOffset, ct)) return true;
                } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException) {
                    // Resolve the row before reporting a provisional style, string or formula error.
                    // The ordinary parser validates the same cells again after rollback.
                }

                position = originalPosition;
                _minimumCellRow = minimumCellRow;
                _maximumCellRow = maximumCellRow;
                _minimumCellColumn = minimumCellColumn;
                _maximumCellColumn = maximumCellColumn;
                return false;
            }

            private bool TryIndexCanonicalImplicitDenseRowCore(
                ref int position,
                int rowIndex,
                int rowOffset,
                CancellationToken ct) {
                int columnIndex = 1;
                int cellsUntilCancellationCheck = 0;
                Utf8ImplicitCellTagCacheEntry[]? tagCache = _implicitCellTagCache;
                while (position < _length) {
                    SkipWhitespaceInsideRow(ref position);
                    if (position <= _length - 6
                        && _buffer![position + 1] == (byte)'/'
                        && TryConsumeCompactRowEnd(ref position)) {
                        // Every indexed cell owns its metadata. Only absent trailing
                        // cells need missing sentinels when a complete row is proven.
                        int indexedFieldCount = columnIndex - 1;
                        if (indexedFieldCount < _fieldCount) {
                            InitializeMetadataRow(rowOffset, indexedFieldCount);
                        }
                        UpdateUsedBounds(rowIndex, columnIndex - 1);
                        // The first row commonly holds headers. All slots must have
                        // qualified tags before a later complete row can reuse them.
                        _cachedImplicitRowReady = !_cachedImplicitRowRejected && _rowCount > 0
                            && _firstColumn == 1 && columnIndex - 1 == _fieldCount;
                        return true;
                    }

                    if (cellsUntilCancellationCheck-- == 0) {
                        ct.ThrowIfCancellationRequested();
                        cellsUntilCancellationCheck = 256;
                    }
                    if (columnIndex == 1 && _minimumCellColumn > 1) {
                        _minimumCellColumn = 1;
                    }
                    int ordinal = columnIndex - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        return false;
                    }

                    tagCache ??= RentImplicitCellTagCache();
                    ref Utf8ImplicitCellTagCacheEntry cachedTag = ref tagCache[ordinal];
                    int cellStart = position;
                    bool dateStylesEnabled = TreatDatesForIndex;
                    bool repeatedTag = cachedTag.Length > 0
                        && cachedTag.DateStylesEnabled == dateStylesEnabled
                        && position <= _length - cachedTag.Length
                        && _buffer!.AsSpan(position, cachedTag.Length).SequenceEqual(
                            _buffer.AsSpan(cachedTag.Start, cachedTag.Length));
                    Utf8CellKind kind;
                    int styleIndex;
                    bool isEmpty;
                    if (repeatedTag) {
                        kind = cachedTag.Kind;
                        styleIndex = cachedTag.StyleIndex;
                        isEmpty = cachedTag.IsEmpty;
                        position += cachedTag.Length;
                    } else if (!TryReadCanonicalImplicitCellAttributes(
                                   ref position, out kind, out styleIndex, out isEmpty)) {
                        return false;
                    }
                    int tagLength = position - cellStart;

                    int cellIndex = rowOffset + ordinal;
                    bool sharedFormulaFollower = false;
                    bool indexedFormula = false;
                    bool generalCell = false;
                    bool hasCachedValue = false;
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!isEmpty) {
                        int contentPosition = position;
                        if (!TryIndexCompactValueCell(
                                ref position, cellIndex, kind,
                                out hasCachedValue, out valueStart, out valueLength)) {
                            position = contentPosition;
                            // The formula and general indexers write only the parts the cell contains.
                            _valueStarts![cellIndex] = 0;
                            _valueLengths![cellIndex] = -1;
                            if (_formulaStarts != null) {
                                _formulaStarts[cellIndex] = 0;
                                _formulaLengths![cellIndex] = -1;
                            }
                            if (!TryIndexCompactFormulaCell(ref position, cellIndex, kind,
                                    out hasCachedValue, out valueStart, out valueLength)) {
                                // Other element content uses the general cell indexer and keeps the full XML check.
                                _sheetDataSupportsFastValidation = false;
                                position = contentPosition;
                                if (!TryIndexCell(ref position, cellIndex, kind,
                                        out sharedFormulaFollower, out hasCachedValue, out valueStart, out valueLength)) {
                                    return false;
                                }
                            }
                            indexedFormula = _formulaLengths != null && _formulaLengths[cellIndex] >= 0;
                            generalCell = true;
                        }
                    }

                    // Exact cached tags already proved their style against this
                    // reader's immutable StylesCache. Shared-string indexes belong
                    // to each value and must still be validated on every cell.
                    int sharedStringIndex = repeatedTag && kind != Utf8CellKind.SharedString && !sharedFormulaFollower
                        ? -1
                        : ValidateIndexedCell(
                            rowIndex, columnIndex, kind, styleIndex,
                            sharedFormulaFollower, hasCachedValue, valueStart, valueLength);
                    if (sharedStringIndex >= 0) {
                        _valueStarts![cellIndex] = sharedStringIndex;
                        _valueLengths![cellIndex] = SharedStringIndexValueLength;
                    } else if (valueLength < 0 && !generalCell) {
                        // The general cell indexer stores inline text itself; only compact cells are reset here.
                        _valueStarts![cellIndex] = 0;
                        _valueLengths![cellIndex] = -1;
                    }
                    if (_formulaStarts != null && !indexedFormula) {
                        _formulaStarts[cellIndex] = 0;
                        _formulaLengths![cellIndex] = -1;
                    }
                    byte encodedKind = repeatedTag ? cachedTag.EncodedKind : EncodeCellKind(kind, styleIndex);
                    _cellKinds![cellIndex] = encodedKind;
                    if (!repeatedTag) {
                        cachedTag = new Utf8ImplicitCellTagCacheEntry(
                            cellStart, tagLength, kind, styleIndex, isEmpty, encodedKind, dateStylesEnabled);
                    }
                    columnIndex++;
                }

                return false;
            }
        }
    }
}
