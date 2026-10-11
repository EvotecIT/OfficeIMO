using System.Data.Common;
using System.Globalization;

namespace OfficeIMO.Odbc;

internal static unsafe class NativeValue {
    internal static short SqlType(DbDataReader reader, int ordinal) {
        Type managed = reader.GetFieldType(ordinal);
        if (managed == typeof(short)) return 5;
        if (managed == typeof(int)) return 4;
        string type = reader.GetDataTypeName(ordinal).ToUpperInvariant();
        if (type.Contains("BOOL", StringComparison.Ordinal)) return -7;
        if (type.Contains("DATETIME", StringComparison.Ordinal) || type.Contains("TIMESTAMP", StringComparison.Ordinal)) return 93;
        if (type.Contains("INT", StringComparison.Ordinal)) return -5;
        if (type.Contains("REAL", StringComparison.Ordinal) || type.Contains("DOUBLE", StringComparison.Ordinal) || type.Contains("FLOAT", StringComparison.Ordinal)) return 8;
        return managed == typeof(double) || managed == typeof(float) ? (short)8 :
            managed == typeof(int) || managed == typeof(long) || managed == typeof(short) ? (short)-5 :
            managed == typeof(bool) ? (short)-7 : managed == typeof(DateTime) ? (short)93 :
            managed == typeof(byte[]) ? (short)-3 : (short)-9;
    }

    internal static nuint ColumnSize(short sqlType) => sqlType switch { 8 => 15, -5 => 19, 4 => 10, 5 => 5, -7 => 1, 93 => 27, _ => 32 * 1024 * 1024 };
    internal static short Precision(short sqlType) => sqlType switch { 8 => 15, -5 => 19, 4 => 10, 5 => 5, -7 => 1, 93 => 7, _ => 0 };
    internal static nint OctetLength(short sqlType) => sqlType switch { -9 => 64 * 1024 * 1024, -7 => 1, 5 => 2, 4 => 4, 93 => 16, _ => 8 };
    internal static string TypeName(short sqlType) => sqlType switch { 8 => "DOUBLE", -5 => "BIGINT", 4 => "INTEGER", 5 => "SMALLINT", -7 => "BIT", 93 => "TIMESTAMP", -3 => "VARBINARY", _ => "WVARCHAR" };

    internal static short Write(StatementHandle statement, ushort column, short targetType, nint buffer, nint capacity, nint indicator, bool advance) {
        if (statement.Reader is not { } reader || !statement.OnRow) return statement.Fail("24000", "Fetch a row before reading column data.");
        if (column == 0 || column > reader.FieldCount) return statement.Fail("07009", "Invalid column ordinal.");
        if (capacity < 0) return statement.Fail("HY090", "Negative buffer length.");
        if (advance && statement.DataColumn != column) {
            statement.DataColumn = column;
            statement.DataOffset = 0;
        }
        if (advance && statement.DataOffset == -1) return 100;
        object value = reader.GetValue(column - 1);
        if (value is DBNull) {
            if (indicator == 0) return statement.Fail("22002", "A null value requires an indicator buffer.");
            *(nint*)indicator = -1;
            if (advance) statement.DataOffset = -1;
            return 0;
        }
        if (targetType == 99) targetType = SqlType(reader, column - 1) switch { 8 => 8, 4 => 4, 5 => 5, -5 => -25, -7 => -7, 93 => 93, -3 => -2, _ => -8 };
        try {
            if (targetType is 1 or -8 or -2) {
                bool wide = targetType == -8;
                byte[] bytes = targetType == -2 && value is byte[] binary ? binary : NativeText.Bytes(Text(value), wide);
                int offset = advance ? statement.DataOffset : 0;
                if (offset >= bytes.Length && offset > 0) return 100;
                int remaining = bytes.Length - offset;
                if (indicator != 0) *(nint*)indicator = remaining;
                if (buffer == 0 || capacity == 0) return remaining > 0 ? (short)1 : (short)0;
                int terminator = targetType == -2 ? 0 : wide ? 2 : 1;
                if (capacity < terminator) return statement.Fail("HY090", "Buffer cannot hold a string terminator.");
                int copied = (int)Math.Min(remaining, Math.Min(int.MaxValue, capacity - terminator));
                if (wide) copied -= copied % 2;
                bytes.AsSpan(offset, copied).CopyTo(new Span<byte>((void*)buffer, copied));
                if (terminator != 0) new Span<byte>((byte*)buffer + copied, terminator).Clear();
                if (advance) statement.DataOffset = copied == remaining ? -1 : offset + copied;
                if (copied < remaining) {
                    statement.Diagnostics.Add(new("01004", "Column data was truncated; call SQLGetData again for the remainder."));
                    return 1;
                }
                return 0;
            }
            if (buffer == 0) return statement.Fail("HY009", "A numeric or temporal result requires a value buffer.");
            bool fractional = false;
            object integral = targetType is 5 or -15 or -17 or 4 or -16 or -18 or -25 or -27
                ? Truncate(value, out fractional) : value;
            int size;
            switch (targetType) {
                case -7:
                    double bit = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    if (!double.IsFinite(bit) || bit < 0 || bit >= 2) throw new OverflowException();
                    fractional = bit != Math.Truncate(bit);
                    *(byte*)buffer = (byte)Math.Truncate(bit); size = 1; break;
                case 5: case -15: *(short*)buffer = Convert.ToInt16(integral, CultureInfo.InvariantCulture); size = 2; break;
                case -17: *(ushort*)buffer = Convert.ToUInt16(integral, CultureInfo.InvariantCulture); size = 2; break;
                case 4: case -16: *(int*)buffer = Convert.ToInt32(integral, CultureInfo.InvariantCulture); size = 4; break;
                case -18: *(uint*)buffer = Convert.ToUInt32(integral, CultureInfo.InvariantCulture); size = 4; break;
                case -25: *(long*)buffer = Convert.ToInt64(integral, CultureInfo.InvariantCulture); size = 8; break;
                case -27: *(ulong*)buffer = Convert.ToUInt64(integral, CultureInfo.InvariantCulture); size = 8; break;
                case 7:
                    float single = Convert.ToSingle(value, CultureInfo.InvariantCulture);
                    if (!float.IsFinite(single)) throw new OverflowException();
                    *(float*)buffer = single; size = 4; break;
                case 8: *(double*)buffer = Convert.ToDouble(value, CultureInfo.InvariantCulture); size = 8; break;
                case 91: case 92: case 93:
                    DateTime date = value is DateTime typed ? typed : DateTime.Parse(Text(value), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                    ushort* parts = (ushort*)buffer;
                    if (targetType == 92) {
                        fractional = date.Ticks % TimeSpan.TicksPerSecond != 0;
                        parts[0] = (ushort)date.Hour; parts[1] = (ushort)date.Minute; parts[2] = (ushort)date.Second; size = 6; break;
                    }
                    if (targetType == 91) fractional = date.TimeOfDay != TimeSpan.Zero;
                    parts[0] = (ushort)date.Year; parts[1] = (ushort)date.Month; parts[2] = (ushort)date.Day;
                    size = 6;
                    if (targetType == 93) {
                        parts[3] = (ushort)date.Hour; parts[4] = (ushort)date.Minute; parts[5] = (ushort)date.Second;
                        *(uint*)((byte*)buffer + 12) = (uint)(date.Ticks % TimeSpan.TicksPerSecond * 100);
                        size = 16;
                    }
                    break;
                default: return statement.Fail("HYC00", "Unsupported C result type.");
            }
            if (indicator != 0) *(nint*)indicator = size;
            if (advance) statement.DataOffset = -1;
            if (fractional) {
                statement.Diagnostics.Add(new("01S07", "Value lost precision during conversion."));
                return 1;
            }
            return 0;
        } catch (OverflowException) { return statement.Fail("22003", "Numeric value is out of range."); }
        catch (FormatException) { return statement.Fail("22018", "Value cannot be converted to the requested type."); }
        catch (InvalidCastException) { return statement.Fail("22018", "Value cannot be converted to the requested type."); }
    }

    private static object Truncate(object value, out bool fractional) {
        object result = value switch {
            double number => Math.Truncate(number), float number => MathF.Truncate(number),
            decimal number => decimal.Truncate(number), _ => value
        };
        fractional = !result.Equals(value);
        return result;
    }

    internal static object? Read(ParameterBinding binding) {
        nint length = binding.Indicator != 0 ? *(nint*)binding.Indicator : binding.Length;
        if (length == -1) return DBNull.Value;
        if (length < 0 && length != -3) throw new NotSupportedException("Data-at-execution parameters are not supported.");
        if (binding.Buffer == 0) throw new ArgumentException("Parameter value buffer is required.");
        return binding.Type switch {
            1 => NativeText.Read(binding.Buffer, checked((int)length), false),
            -8 => NativeText.Read(binding.Buffer, length == -3 ? -3 : checked((int)length / 2), true),
            -7 => *(byte*)binding.Buffer != 0,
            5 or -15 => *(short*)binding.Buffer,
            -17 => *(ushort*)binding.Buffer,
            4 or -16 => *(int*)binding.Buffer,
            -18 => *(uint*)binding.Buffer,
            -25 => *(long*)binding.Buffer,
            7 => *(float*)binding.Buffer,
            8 => *(double*)binding.Buffer,
            91 or 93 => ReadDate(binding.Buffer, binding.Type == 93),
            _ => throw new NotSupportedException("Unsupported C parameter type.")
        };
    }

    private static DateTime ReadDate(nint buffer, bool timestamp) {
        ushort* parts = (ushort*)buffer;
        DateTime result = new(parts[0], parts[1], parts[2], timestamp ? parts[3] : 0, timestamp ? parts[4] : 0, timestamp ? parts[5] : 0);
        return timestamp ? result.AddTicks(*(uint*)((byte*)buffer + 12) / 100) : result;
    }

    private static string Text(object value) => value is DateTime date ? date.ToString("O", CultureInfo.InvariantCulture) :
        Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
}
