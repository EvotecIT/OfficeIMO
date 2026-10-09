#nullable enable

using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        // Scalar access and lazy materialization share the same per-row value cache.
        private sealed partial class ExcelXmlRangeDataReader {
            /// <inheritdoc />
            public override bool GetBoolean(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Boolean);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Boolean) {
                    return _currentBooleanValues[ordinal];
                }

                object value = GetNonDbNullValue(ordinal);
                return value is bool boolean ? boolean : Convert.ToBoolean(value, _culture);
            }

            /// <inheritdoc />
            public override byte GetByte(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToByte(value)
                : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? Convert.ToByte(decimalValue)
                : Convert.ToByte(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
                throw new NotSupportedException("Excel range fields are exposed as scalar values.");

            /// <inheritdoc />
            public override char GetChar(int ordinal) => Convert.ToChar(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) {
                string value = Convert.ToString(GetValue(ordinal), _culture) ?? string.Empty;
                if (buffer == null) {
                    return value.Length;
                }

                if (dataOffset >= value.Length || length == 0) {
                    return 0;
                }

                int offset = (int)dataOffset;
                int count = Math.Min(length, value.Length - offset);
                if (count <= 0) {
                    return 0;
                }

                value.CopyTo(offset, buffer, bufferOffset, count);
                return count;
            }

            /// <inheritdoc />
            public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

            /// <inheritdoc />
            public override DateTime GetDateTime(int ordinal) {
                if (TryGetUnloadedDateTime(ordinal, out DateTime indexedDate)) return indexedDate;
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.DateTime);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.DateTime) {
                    return _currentDateTimeValues[ordinal];
                }
                if (IsCurrentDateSerial(ordinal)) {
                    return MaterializeDateSerial(ordinal);
                }

                object value = GetNonDbNullValue(ordinal);
                return value is DateTime dateTime ? dateTime : Convert.ToDateTime(value, _culture);
            }

            /// <inheritdoc />
            public override decimal GetDecimal(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToDecimal(value)
                : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? decimalValue
                : Convert.ToDecimal(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override double GetDouble(int ordinal) {
                return TryGetPrimitiveDouble(ordinal, out double value)
                    ? value
                    : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? (double)decimalValue
                    : Convert.ToDouble(GetNonDbNullValue(ordinal), _culture);
            }

            /// <inheritdoc />
            [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Excel reader column types are closed scalar conversion tokens; OfficeIMO never activates or reflects over their public members.")]
            [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
            public override Type GetFieldType(int ordinal) => _columnTypes[ordinal];

            /// <inheritdoc />
            public override float GetFloat(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? (float)value
                : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? (float)decimalValue
                : Convert.ToSingle(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override Guid GetGuid(int ordinal) {
                object value = GetNonDbNullValue(ordinal);
                return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _culture)!);
            }

            /// <inheritdoc />
            public override short GetInt16(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToInt16(value)
                : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? Convert.ToInt16(decimalValue)
                : Convert.ToInt16(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override int GetInt32(int ordinal) {
                if (TryGetUnloadedInt32(ordinal, out int integer)) return integer;
                return TryGetPrimitiveDouble(ordinal, out double value)
                    ? ConvertDataReaderInt32(value)
                    : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? Convert.ToInt32(decimalValue)
                    : ConvertDataReaderInt32(GetNonDbNullValue(ordinal), _culture);
            }

            /// <inheritdoc />
            public override long GetInt64(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToInt64(value)
                : TryGetCachedDecimal(ordinal, out decimal decimalValue) ? Convert.ToInt64(decimalValue)
                : Convert.ToInt64(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override string GetName(int ordinal) => _columnNames[ordinal];

            /// <inheritdoc />
            public override int GetOrdinal(string name) {
                _ordinals ??= CreateOrdinalMap(_columnNames);
                if (_ordinals.TryGetValue(name, out int ordinal)) {
                    return ordinal;
                }

                throw new IndexOutOfRangeException(name);
            }

            /// <inheritdoc />
            public override string GetString(int ordinal) {
                if (TryGetUnloadedString(ordinal, out string indexedText)) return indexedText;
                object value = GetNonDbNullValue(ordinal);
                return value is string text ? text : Convert.ToString(value, _culture) ?? string.Empty;
            }

            /// <inheritdoc />
            public override object GetValue(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal);
                object? value = MaterializeCurrentValue(ordinal);
                return ToDataReaderValue(value);
            }

            /// <inheritdoc />
            public override int GetValues(object[] values) {
                EnsureOpenRow();
                MaterializeAllCurrentRowValues();
                MaterializeAllPrimitiveCurrentValues();
                return CopyDataReaderValues(_currentRow!, _fieldCount, values);
            }

            /// <inheritdoc />
            public override bool IsDBNull(int ordinal) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (_currentRowIsBlank || _currentRow == null) {
                    return true;
                }
                if (!IsCurrentStreamingRow) {
                    return _currentRow[ordinal] == null || _currentRow[ordinal] == DBNull.Value;
                }
                if (_currentValueLoaded[ordinal]) {
                    return _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None
                        && (_currentRow[ordinal] == null || _currentRow[ordinal] == DBNull.Value);
                }
                if (_utf8Source != null) {
                    return _utf8Source.IsNull(ordinal + _utf8SourceOrdinalOffset);
                }

                EnsureCurrentValue(ordinal);
                return _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None
                    && (_currentRow[ordinal] == null || _currentRow[ordinal] == DBNull.Value);
            }

            private object GetNonDbNullValue(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal);
                object? value = MaterializeCurrentValue(ordinal);
                if (value == null || value == DBNull.Value) {
                    throw new InvalidCastException($"Column '{GetName(ordinal)}' contains DBNull.");
                }

                return value is ExcelDataReaderDateSerial dateSerial ? dateSerial.Materialize() : value;
            }

            private bool TryGetPrimitiveDouble(int ordinal, out double value) {
                if (TryGetUnloadedNumber(ordinal, out value)) return true;
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Numeric);
                if (IsCurrentStreamingRow
                    && (_currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Double || IsCurrentDateSerial(ordinal))) {
                    value = _currentDoubleValues[ordinal];
                    return true;
                }
                if (_currentRow![ordinal] is ExcelDataReaderDateSerial dateSerial) {
                    value = dateSerial.Serial;
                    return true;
                }
                if (_utf8Source != null && (_currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.DateTime || _currentRow[ordinal] is DateTime)) {
                    _utf8Source.ReadValue(ordinal + _utf8SourceOrdinalOffset, XmlDataReaderTargetKind.Numeric,
                        out XmlDataReaderPrimitiveKind kind, out value, out _, out _, out _, out _, out _);
                    if (kind == XmlDataReaderPrimitiveKind.Double) return true;
                }

                value = 0;
                return false;
            }

            // Called after numeric access has populated the current value cache.
            private bool TryGetCachedDecimal(int ordinal, out decimal value) {
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Decimal) {
                    value = _currentDecimalValues[ordinal];
                    return true;
                }
                value = default;
                return false;
            }

            private bool IsCurrentDateSerial(int ordinal) => IsCurrentStreamingRow
                && (_currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.CalendarDateSerial
                    || _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.ElapsedDateSerial);

            private DateTime MaterializeDateSerial(int ordinal) => _owner.FromExcelSerialDate(
                _currentDoubleValues[ordinal], _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.CalendarDateSerial);

            private object? MaterializeCurrentValue(int ordinal) {
                if (!IsCurrentStreamingRow || _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None) {
                    return _currentRow![ordinal];
                }

                object value = _currentValues[ordinal] ?? (_currentPrimitiveKinds[ordinal] switch {
                    XmlDataReaderPrimitiveKind.Double => _currentDoubleValues[ordinal],
                    XmlDataReaderPrimitiveKind.Decimal => _currentDecimalValues[ordinal],
                    XmlDataReaderPrimitiveKind.CalendarDateSerial or XmlDataReaderPrimitiveKind.ElapsedDateSerial => MaterializeDateSerial(ordinal),
                    XmlDataReaderPrimitiveKind.DateTime => _currentDateTimeValues[ordinal],
                    XmlDataReaderPrimitiveKind.Boolean => BoxBoolean(_currentBooleanValues[ordinal]),
                    _ => _currentRow![ordinal]!
                });
                _currentValues[ordinal] = value;
                // Retain the serial alongside the boxed date so later numeric getters
                // still return the workbook's value, including the 1904 date system.
                if (!IsCurrentDateSerial(ordinal)) _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.None;
                return value;
            }

            private void MaterializeAllPrimitiveCurrentValues() {
                if (!IsCurrentStreamingRow) {
                    return;
                }

                for (int i = 0; i < _currentPrimitiveKinds.Length; i++) {
                    if (_currentPrimitiveKinds[i] != XmlDataReaderPrimitiveKind.None) {
                        _ = MaterializeCurrentValue(i);
                    }
                }
            }

        }
    }
}
