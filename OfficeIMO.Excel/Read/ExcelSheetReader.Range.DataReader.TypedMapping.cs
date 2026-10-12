namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
            bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

            bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) throw new IndexOutOfRangeException();
                if (_utf8Source == null || !IsCurrentStreamingRow || _currentRowIsBlank
                    || _owner._opt.NumericAsDecimal || _owner._opt.CellValueConverter != null
                    || !_utf8Source.CanUseTypedMappingGetter(ordinal + _utf8SourceOrdinalOffset, targetType)) {
                    return false;
                }
                if (targetType != typeof(DateTime)) return true;

                try {
                    // Qualification and the typed getter share the existing primitive cache.
                    // Keep previously materialized values/serials in their original cache state.
                    return _currentValueLoaded[ordinal]
                        ? _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.DateTime
                            || _utf8Source.TryGetDateTime(ordinal + _utf8SourceOrdinalOffset, out _)
                        : TryGetUnloadedDateTime(ordinal, out _);
                } catch (ArgumentException) {
                    // Generic mapping constructs the model before reporting an invalid date serial.
                    return false;
                }
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
                // The data reader validates numeric tokens and caches successful date conversion.
                if (targetType == typeof(DateTime)) return dateStyle;
                // Date-style numeric projection retains the existing serial/converter/range-validation fallback.
                return !dateStyle && (targetType == typeof(byte) || targetType == typeof(short)
                    || targetType == typeof(int) || targetType == typeof(long)
                    || targetType == typeof(float) || targetType == typeof(double)
                    || targetType == typeof(decimal));
            }
        }
    }
}
