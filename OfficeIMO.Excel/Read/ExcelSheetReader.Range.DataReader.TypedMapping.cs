namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
            bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

            bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) throw new IndexOutOfRangeException();
                return _utf8Source != null && IsCurrentStreamingRow && !_currentRowIsBlank
                    && !_owner._opt.NumericAsDecimal && _owner._opt.CellValueConverter == null
                    && _utf8Source.CanUseTypedMappingGetter(ordinal + _utf8SourceOrdinalOffset, targetType);
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>Qualifies native cell kinds without materializing boxed values or changing text conversion.</summary>
            internal bool CanUseTypedMappingGetter(int ordinal, Type targetType) {
                EnsureNotDisposed();
                int cellIndex = _currentRowOffset + ordinal;
                if (_formulaLengths != null && _formulaLengths[cellIndex] >= 0) return false;
                byte encodedKind = _cellKinds![cellIndex];
                Utf8CellKind kind = (Utf8CellKind)(encodedKind & CellKindMask);
                if (targetType == typeof(string)) {
                    return kind is Utf8CellKind.String or Utf8CellKind.InlineString
                        or Utf8CellKind.SharedString or Utf8CellKind.Error;
                }
                if (targetType == typeof(bool)) return kind == Utf8CellKind.Boolean;
                if (kind != Utf8CellKind.Number) return false;
                bool dateStyle = (encodedKind & DateStyleCellKindFlag) != 0;
                if (targetType == typeof(DateTime)) {
                    // A malformed numeric token can fall back to source text. Preserve its RoundtripKind conversion.
                    if (!dateStyle) return false;
                    try {
                        return TryGetDateTime(ordinal, out _);
                    } catch (ArgumentException) {
                        // Generic mapping constructs the model before reporting an invalid date serial.
                        return false;
                    }
                }
                // Date-style numeric projection retains the existing serial/converter/range-validation fallback.
                return !dateStyle && (targetType == typeof(byte) || targetType == typeof(short)
                    || targetType == typeof(int) || targetType == typeof(long)
                    || targetType == typeof(float) || targetType == typeof(double)
                    || targetType == typeof(decimal));
            }
        }
    }
}
