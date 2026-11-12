namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
            bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

            bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) =>
                CanUseTypedMappingGetter(ordinal, targetType, requireNonNull: false);

            bool IDataReaderTypedMappingCompatibility.CanUseNonNullTypedGetter(int ordinal, Type targetType) =>
                CanUseTypedMappingGetter(ordinal, targetType, requireNonNull: true);

            private bool CanUseTypedMappingGetter(int ordinal, Type targetType, bool requireNonNull) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) throw new IndexOutOfRangeException();
                if (_utf8Source == null || !IsCurrentStreamingRow || _currentRowIsBlank
                    || _owner._opt.CellValueConverter != null
                    || _owner._opt.NumericAsDecimal && !CanUseTypedGetterWithDecimalNumbers(ordinal, targetType)
                    || !_utf8Source.CanUseTypedMappingGetter(
                        ordinal + _utf8SourceOrdinalOffset, targetType, requireNonNull)) {
                    return false;
                }
                if (requireNonNull && _currentValueLoaded[ordinal]
                    && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None
                    && (_currentRow![ordinal] == null || _currentRow[ordinal] == DBNull.Value)) {
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

            /// <summary>
            /// With <see cref="ExcelReaderOptions.NumericAsDecimal"/>, numbers are decoded as decimals.
            /// A typed getter qualifies only where it returns what conversion of that decimal value
            /// returns: decimals read the same decoded value, integers must be exact integer tokens,
            /// and text, Boolean and date cells do not depend on the numeric representation.
            /// </summary>
            private bool CanUseTypedGetterWithDecimalNumbers(int ordinal, Type targetType) {
                if (targetType == typeof(decimal) || targetType == typeof(string)
                    || targetType == typeof(bool) || targetType == typeof(DateTime)) {
                    return true;
                }

                return targetType == typeof(int)
                    && _utf8Source!.TryGetExactInt32(ordinal + _utf8SourceOrdinalOffset, out _);
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>Qualifies native cell kinds without materializing boxed values or changing text conversion.</summary>
            internal bool CanUseTypedMappingGetter(int ordinal, Type targetType, bool requireNonNull) {
                EnsureNotDisposed();
                int cellIndex = _currentRowOffset + ordinal;
                if (_formulaLengths != null && _formulaLengths[cellIndex] >= 0) return false;
                byte encodedKind = _cellKinds![cellIndex];
                Utf8CellKind kind = (Utf8CellKind)(encodedKind & CellKindMask);
                if (requireNonNull && _valueLengths![cellIndex] < 0
                    && !(kind == Utf8CellKind.SharedString
                        && _valueLengths[cellIndex] == SharedStringIndexValueLength)) {
                    return false;
                }
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
