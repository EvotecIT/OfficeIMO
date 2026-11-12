using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader : IDataReaderFastMappingValues, IDataReaderTypedRowQualification {
            bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

            object? IDataReaderTypedRowQualification.PrepareNonNullTypedRowQualification(int[] ordinals, Type[] targetTypes) {
                if (_utf8Source == null || ordinals.Length == 0 || ordinals.Length != targetTypes.Length) return null;
                var targets = new byte[ordinals.Length];
                for (int index = 0; index < ordinals.Length; index++) {
                    if ((uint)ordinals[index] >= (uint)_fieldCount) return null;
                    targets[index] = ExcelUtf8RangeRowSource.GetTypedMappingTarget(targetTypes[index]);
                }
                return new TypedRowQualification(this, ordinals, targetTypes, targets);
            }

            bool IDataReaderTypedRowQualification.CanUseNonNullTypedGetters(object preparedState) {
                var prepared = (TypedRowQualification)preparedState;
                if (!ReferenceEquals(prepared.Reader, this)) throw new ArgumentException("The qualification belongs to another reader.", nameof(preparedState));
                int[] ordinals = prepared.Ordinals;
                Type[] targetTypes = prepared.TargetTypes;
                EnsureOpenRow();
                ExcelUtf8RangeRowSource source = _utf8Source!;
                if (!IsCurrentStreamingRow || _currentRowIsBlank || _owner._opt.CellValueConverter != null) return false;
                if (_owner._opt.NumericAsDecimal) {
                    for (int index = 0; index < ordinals.Length; index++) {
                        if (!CanUseTypedMappingGetter(ordinals[index], targetTypes[index], requireNonNull: true)) return false;
                    }
                    return true;
                }

                byte[] targets = prepared.Targets;
                bool[] loaded = _currentValueLoaded;
                for (int index = 0; index < ordinals.Length; index++) {
                    int ordinal = ordinals[index];
                    // Values a caller already read keep the cache-aware per-field rules.
                    int result = loaded[ordinal]
                        ? ExcelUtf8RangeRowSource.TypedMappingNeedsFieldCheck
                        : source.QualifyUnloadedNonNullTypedMapping(ordinal + _utf8SourceOrdinalOffset, targets[index]);
                    if (result == ExcelUtf8RangeRowSource.TypedMappingRejected) return false;
                    if (result == ExcelUtf8RangeRowSource.TypedMappingNeedsFieldCheck
                        && !CanUseTypedMappingGetter(ordinal, targetTypes[index], requireNonNull: true)) {
                        return false;
                    }
                }
                return true;
            }

            /// <summary>Target kinds of the mapped fields, resolved once per row reader.</summary>
            private sealed class TypedRowQualification {
                internal TypedRowQualification(ExcelXmlRangeDataReader reader, int[] ordinals, Type[] targetTypes, byte[] targets) {
                    Reader = reader;
                    Ordinals = ordinals;
                    TargetTypes = targetTypes;
                    Targets = targets;
                }

                internal ExcelXmlRangeDataReader Reader { get; }
                internal int[] Ordinals { get; }
                internal Type[] TargetTypes { get; }
                internal byte[] Targets { get; }
            }

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
            internal const int TypedMappingRejected = 0;
            internal const int TypedMappingAccepted = 1;
            internal const int TypedMappingNeedsFieldCheck = 2;
            private const byte TargetOther = 0;
            private const byte TargetString = 1;
            private const byte TargetBoolean = 2;
            private const byte TargetNumber = 3;
            private const byte TargetDateTime = 4;

            internal static byte GetTypedMappingTarget(Type targetType) {
                if (targetType == typeof(string)) return TargetString;
                if (targetType == typeof(bool)) return TargetBoolean;
                if (targetType == typeof(DateTime)) return TargetDateTime;
                return targetType == typeof(byte) || targetType == typeof(short)
                    || targetType == typeof(int) || targetType == typeof(long)
                    || targetType == typeof(float) || targetType == typeof(double)
                    || targetType == typeof(decimal) ? TargetNumber : TargetOther;
            }

            /// <summary>
            /// Applies the non-null <see cref="CanUseTypedMappingGetter"/> rules to a field the row has
            /// not loaded. A date cell is accepted here only when indexing already converted it; other
            /// date cells return <see cref="TypedMappingNeedsFieldCheck"/> for the converting check.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            internal int QualifyUnloadedNonNullTypedMapping(int ordinal, byte target) {
                EnsureNotDisposed();
                int cellIndex = _currentRowOffset + ordinal;
                if (_formulaLengths != null && _formulaLengths[cellIndex] >= 0) return TypedMappingRejected;
                byte encodedKind = _cellKinds![cellIndex];
                Utf8CellKind kind = (Utf8CellKind)(encodedKind & CellKindMask);
                int valueLength = _valueLengths![cellIndex];
                if (valueLength < 0 && !(kind == Utf8CellKind.SharedString && valueLength == SharedStringIndexValueLength)) {
                    return TypedMappingRejected;
                }
                bool dateStyle = (encodedKind & DateStyleCellKindFlag) != 0;
                switch (target) {
                    case TargetString:
                        return kind is Utf8CellKind.String or Utf8CellKind.InlineString
                            or Utf8CellKind.SharedString or Utf8CellKind.Error ? TypedMappingAccepted : TypedMappingRejected;
                    case TargetBoolean:
                        return kind == Utf8CellKind.Boolean ? TypedMappingAccepted : TypedMappingRejected;
                    case TargetNumber:
                        return kind == Utf8CellKind.Number && !dateStyle ? TypedMappingAccepted : TypedMappingRejected;
                    case TargetDateTime:
                        if (kind != Utf8CellKind.Number || !dateStyle) return TypedMappingRejected;
                        byte[]? tags = _decodedTags;
                        return tags != null && (tags[cellIndex] & DecodedKindMask) == DecodedDate
                            ? TypedMappingAccepted : TypedMappingNeedsFieldCheck;
                    default:
                        return TypedMappingRejected;
                }
            }

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
