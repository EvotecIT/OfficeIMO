using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            private void CompleteCurrentXmlRow() {
                if (!_currentRowActive || _currentRowFinished) return;

                // Complete the character-bounded row before exposing any scalar:
                // duplicate or descending cells must agree with bulk getters.
                ThrowIfReadCancellationRequested();
                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();
                    if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _currentRowDepth && _reader.LocalName == "row") {
                        _currentRowActive = false;
                        _currentRowFinished = true;
                        break;
                    }
                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "c") continue;

                    int columnIndex = GetXmlCellColumnIndex(_reader, ref _currentNextCellColumnIndex);
                    if (columnIndex < _firstColumn || columnIndex > _lastColumn) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }
                    int ordinal = columnIndex - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }

                    if (_currentCellPresence != null) _currentCellPresence[ordinal] = true;
                    _currentValues[ordinal] = ReadBoundedXmlCellValue();
                    _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.None;
                    _currentValueLoaded[ordinal] = true;
                }
            }

            private void ReadBufferedXmlRowValues(object?[] values, bool[]? presence) {
                if (_reader.IsEmptyElement) return;
                int depth = _reader.Depth;
                int nextColumn = 1;
                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();
                    if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == depth && _reader.LocalName == "row") return;
                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "c") continue;

                    int columnIndex = GetXmlCellColumnIndex(_reader, ref nextColumn);
                    if (columnIndex < _firstColumn || columnIndex > _lastColumn) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }
                    int ordinal = columnIndex - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }

                    if (presence != null) presence[ordinal] = true;
                    values[ordinal] = ReadBoundedXmlCellValue();
                }
            }

            private object? ReadBoundedXmlCellValue() {
                try {
                    return _owner.ReadXmlCellValue(_reader, ReadXmlCellTypeAttribute(_reader),
                        preserveDateSerial: true, textBudget: _xmlTextBudget!);
                } catch (InvalidDataException) when (_xmlTextBudget!.WasExceeded) {
                    _xmlTextBufferingFailed = true;
                    ReleaseCurrentValueReferences();
                    ReleaseBufferedValueReferences();
                    _currentRow = null;
                    _currentCellPresence = null;
                    _currentRowActive = false;
                    _currentRowFinished = true;
                    throw;
                }
            }

            private void ThrowIfXmlTextBufferingFailed() {
                if (_xmlTextBufferingFailed) {
                    throw ExcelReadLimitFailure.Create(
                        $"The XML range data reader exceeded {nameof(ExcelReadOptions.MaxXmlDataReaderBufferedCharacters)}. Close the reader.");
                }
            }

            private void ReleaseCurrentValueReferences() {
                if (_currentRow != null && !ReferenceEquals(_currentRow, _currentValues)) {
                    Array.Clear(_currentRow, 0, _currentRow.Length);
                }
                Array.Clear(_currentValues, 0, _currentValues.Length);
                Array.Clear(_currentValueLoaded, 0, _currentValueLoaded.Length);
                Array.Clear(_currentPrimitiveKinds, 0, _currentPrimitiveKinds.Length);
            }

            private void ReleaseBufferedValueReferences() {
                if (_bufferedRows != null) {
                    foreach (object?[] values in _bufferedRows.Values) Array.Clear(values, 0, values.Length);
                    _bufferedRows.Clear();
                }
                _bufferedCellPresence?.Clear();
            }
        }
    }
}
