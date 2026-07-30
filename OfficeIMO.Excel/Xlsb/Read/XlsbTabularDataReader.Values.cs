using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        public override string GetName(int ordinal) {
            ValidateOrdinal(ordinal);
            return _headers[ordinal];
        }

        public override int GetOrdinal(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            if (_ordinals.TryGetValue(name, out int ordinal)) {
                return ordinal;
            }

            throw new IndexOutOfRangeException($"Column '{name}' was not found.");
        }

        [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "XLSB column types are returned as DbDataReader schema tokens; OfficeIMO does not activate or reflect over their public members.")]
        [UnconditionalSuppressMessage("Trimming", "IL2073", Justification = "XLSB column types are returned as DbDataReader schema tokens; OfficeIMO does not activate or reflect over their public members.")]
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
        public override Type GetFieldType(int ordinal) {
            ValidateOrdinal(ordinal);
            if (_options.InferSchema) {
                return _columnTypes[ordinal];
            }

            return typeof(object);
        }

        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

        public override object GetValue(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => _strings[ordinal]!,
                XlsbTabularValueKind.Number => GetNumericValue(_numbers[ordinal]),
                XlsbTabularValueKind.Boolean => _booleans[ordinal],
                XlsbTabularValueKind.Date => ConvertDate(_numbers[ordinal]),
                XlsbTabularValueKind.Custom => _customValues[ordinal] ?? DBNull.Value,
                _ => DBNull.Value
            };
        }

        public override int GetValues(object[] values) {
            if (values == null) {
                throw new ArgumentNullException(nameof(values));
            }

            int count = Math.Min(values.Length, FieldCount);
            for (int index = 0; index < count; index++) {
                values[index] = GetValue(index);
            }

            return count;
        }

        public override bool IsDBNull(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Empty
                   || _kinds[ordinal] == XlsbTabularValueKind.Custom
                   && _customValues[ordinal] == null;
        }

        public override string GetString(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] switch {
                XlsbTabularValueKind.Text or XlsbTabularValueKind.Error => _strings[ordinal]!,
                XlsbTabularValueKind.Number => _numbers[ordinal].ToString("R", _options.Culture),
                XlsbTabularValueKind.Boolean => _booleans[ordinal].ToString(),
                XlsbTabularValueKind.Date => ConvertDate(_numbers[ordinal]).ToString(_options.Culture),
                XlsbTabularValueKind.Custom when _customValues[ordinal] != null =>
                    Convert.ToString(_customValues[ordinal], _options.Culture)
                    ?? string.Empty,
                _ => throw new InvalidCastException("The XLSB cell is blank.")
            };
        }

        public override bool GetBoolean(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            if (_kinds[ordinal] == XlsbTabularValueKind.Boolean) {
                return _booleans[ordinal];
            }

            return Convert.ToBoolean(GetValue(ordinal), _options.Culture);
        }

        public override byte GetByte(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToByte(_numbers[ordinal])
                : Convert.ToByte(GetValue(ordinal), _options.Culture);
        }

        public override short GetInt16(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt16(_numbers[ordinal])
                : Convert.ToInt16(GetValue(ordinal), _options.Culture);
        }

        public override int GetInt32(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt32(_numbers[ordinal])
                : Convert.ToInt32(GetValue(ordinal), _options.Culture);
        }

        public override long GetInt64(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? Convert.ToInt64(_numbers[ordinal])
                : Convert.ToInt64(GetValue(ordinal), _options.Culture);
        }

        public override float GetFloat(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? (float)_numbers[ordinal]
                : Convert.ToSingle(GetValue(ordinal), _options.Culture);
        }

        public override double GetDouble(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? _numbers[ordinal]
                : Convert.ToDouble(GetValue(ordinal), _options.Culture);
        }

        public override decimal GetDecimal(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Number
                ? ConvertExcelNumberToDecimal(_numbers[ordinal])
                : Convert.ToDecimal(GetValue(ordinal), _options.Culture);
        }

        public override DateTime GetDateTime(int ordinal) {
            ValidateReadableOrdinal(ordinal);
            return _kinds[ordinal] == XlsbTabularValueKind.Date
                ? ConvertDate(_numbers[ordinal])
                : Convert.ToDateTime(GetValue(ordinal), _options.Culture);
        }

        public override Guid GetGuid(int ordinal) {
            object value = GetValue(ordinal);
            return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _options.Culture)!);
        }

        public override char GetChar(int ordinal) => Convert.ToChar(GetValue(ordinal), _options.Culture);

        public override long GetBytes(
            int ordinal,
            long dataOffset,
            byte[]? buffer,
            int bufferOffset,
            int length) {
            if (GetValue(ordinal) is not byte[] bytes) {
                throw new InvalidCastException("The XLSB cell does not contain binary data.");
            }

            return CopySegment(bytes, dataOffset, buffer, bufferOffset, length);
        }

        public override long GetChars(
            int ordinal,
            long dataOffset,
            char[]? buffer,
            int bufferOffset,
            int length) {
            char[] characters = GetString(ordinal).ToCharArray();
            return CopySegment(characters, dataOffset, buffer, bufferOffset, length);
        }

        public override IEnumerator GetEnumerator() => new DbEnumerator(this, closeReader: false);

        [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type.TypeInitializer or other Type members.")]
        public override DataTable GetSchemaTable() {
            var table = new DataTable("SchemaTable");
            table.Columns.Add("ColumnName", typeof(string));
            table.Columns.Add("ColumnOrdinal", typeof(int));
            table.Columns.Add("DataType", typeof(Type));
            for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                DataRow row = table.NewRow();
                row["ColumnName"] = GetName(ordinal);
                row["ColumnOrdinal"] = ordinal;
                row["DataType"] = GetFieldType(ordinal);
                table.Rows.Add(row);
            }

            return table;
        }
    }
}
