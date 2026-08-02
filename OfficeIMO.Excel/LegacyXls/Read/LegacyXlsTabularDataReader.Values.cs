using OfficeIMO.Excel.LegacyXls.Projection;
using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        public override string GetName(int ordinal) {
            ValidateOrdinal(ordinal);
            return _headers[ordinal];
        }

        public override int GetOrdinal(string name) {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (_ordinals.TryGetValue(name, out int ordinal)) return ordinal;
            throw new IndexOutOfRangeException($"Column '{name}' was not found.");
        }

        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
        public override Type GetFieldType(int ordinal) {
            ValidateOrdinal(ordinal);
            return typeof(object);
        }

        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

        public override object GetValue(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                ValueKind.Text or ValueKind.Error => _strings[ordinal]!,
                ValueKind.Number => GetNumericValue(_numbers[ordinal]),
                ValueKind.Boolean => _booleans[ordinal],
                ValueKind.Date => ConvertDate(_numbers[ordinal]),
                _ => DBNull.Value
            };
        }

        public override int GetValues(object[] values) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            int count = Math.Min(values.Length, FieldCount);
            for (int index = 0; index < count; index++) values[index] = GetValue(index);
            return count;
        }

        public override bool IsDBNull(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == ValueKind.Empty;
        }

        public override string GetString(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                ValueKind.Text or ValueKind.Error => _strings[ordinal]!,
                ValueKind.Number => _numbers[ordinal].ToString("R", _options.Culture),
                ValueKind.Boolean => _booleans[ordinal].ToString(),
                ValueKind.Date => ConvertDate(_numbers[ordinal]).ToString(_options.Culture),
                _ => throw new InvalidCastException("The XLS cell is blank.")
            };
        }

        public override bool GetBoolean(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == ValueKind.Boolean
                ? _booleans[ordinal]
                : Convert.ToBoolean(GetValue(ordinal), _options.Culture);
        }

        public override byte GetByte(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? Convert.ToByte(_numbers[ordinal])
                : Convert.ToByte(GetValue(ordinal), _options.Culture);
        }

        public override short GetInt16(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? Convert.ToInt16(_numbers[ordinal])
                : Convert.ToInt16(GetValue(ordinal), _options.Culture);
        }

        public override int GetInt32(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? Convert.ToInt32(_numbers[ordinal])
                : Convert.ToInt32(GetValue(ordinal), _options.Culture);
        }

        public override long GetInt64(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? Convert.ToInt64(_numbers[ordinal])
                : Convert.ToInt64(GetValue(ordinal), _options.Culture);
        }

        public override float GetFloat(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? (float)_numbers[ordinal]
                : Convert.ToSingle(GetValue(ordinal), _options.Culture);
        }

        public override double GetDouble(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] is ValueKind.Number or ValueKind.Date
                ? _numbers[ordinal]
                : Convert.ToDouble(GetValue(ordinal), _options.Culture);
        }

        public override decimal GetDecimal(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            if (_kinds[ordinal] is ValueKind.Number or ValueKind.Date) {
                try {
                    return (decimal)_numbers[ordinal];
                } catch (OverflowException) {
                    throw new InvalidCastException($"The XLS numeric value '{_numbers[ordinal]}' cannot be represented as decimal.");
                }
            }
            return Convert.ToDecimal(GetValue(ordinal), _options.Culture);
        }

        public override DateTime GetDateTime(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == ValueKind.Date
                ? ConvertDate(_numbers[ordinal])
                : Convert.ToDateTime(GetValue(ordinal), _options.Culture);
        }

        public override Guid GetGuid(int ordinal) {
            object value = GetValue(ordinal);
            return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _options.Culture)!);
        }

        public override char GetChar(int ordinal) => Convert.ToChar(GetValue(ordinal), _options.Culture);

        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) {
            if (GetValue(ordinal) is not byte[] bytes) throw new InvalidCastException("The XLS cell does not contain binary data.");
            return CopySegment(bytes, dataOffset, buffer, bufferOffset, length);
        }

        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
            CopySegment(GetString(ordinal).ToCharArray(), dataOffset, buffer, bufferOffset, length);

        public override IEnumerator GetEnumerator() => new DbEnumerator(this, closeReader: false);

        [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type members.")]
        public override DataTable GetSchemaTable() =>
            ExcelDataReaderSchemaTable.Create(FieldCount, GetName, GetFieldType);

        private object GetNumericValue(double value) {
            if (_options.NumericAsDecimal && !double.IsNaN(value) && !double.IsInfinity(value)) {
                try {
                    return (decimal)value;
                } catch (OverflowException) {
                    // Preserve a finite value that decimal cannot represent as double.
                }
            }
            return value;
        }

        private DateTime ConvertDate(double serial) {
            if (LegacyXlsDateSerialConverter.TryConvert(serial, _uses1904DateSystem, out DateTime value)) return value;
            throw new InvalidCastException($"The XLS numeric value '{serial}' is not a valid Excel date.");
        }

        private static long CopySegment<T>(T[] source, long dataOffset, T[]? destination, int destinationOffset, int length) {
            if (dataOffset < 0 || dataOffset > source.Length) throw new ArgumentOutOfRangeException(nameof(dataOffset));
            if (destination == null) return source.Length;
            int count = Math.Min(source.Length - checked((int)dataOffset), length);
            Array.Copy(source, checked((int)dataOffset), destination, destinationOffset, count);
            return count;
        }
    }
}
