using System.Data.Common;
using System.Globalization;

namespace OfficeIMO.Odbc {

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

        internal static short DeclaredSqlType(string typeName) {
            string type = typeName.ToUpperInvariant();
            if (type.Contains("BOOL", StringComparison.Ordinal)) return -7;
            if (type.Contains("DATETIME", StringComparison.Ordinal) || type.Contains("TIMESTAMP", StringComparison.Ordinal)) return 93;
            if (type.Contains("INT", StringComparison.Ordinal)) return -5;
            if (type.Contains("CHAR", StringComparison.Ordinal) || type.Contains("CLOB", StringComparison.Ordinal) || type.Contains("TEXT", StringComparison.Ordinal)) return -9;
            if (type.Contains("REAL", StringComparison.Ordinal) || type.Contains("DOUBLE", StringComparison.Ordinal) || type.Contains("FLOAT", StringComparison.Ordinal)) return 8;
            if (type.Contains("BLOB", StringComparison.Ordinal)) return -3;
            // Expressions and dynamic numeric affinities have no value-independent SQL type.
            return 0;
        }

        // SELECT expressions can exceed imported-cell bounds; unknown variable result sizes are not maxima.
        internal static nuint ColumnSize(short sqlType) => sqlType switch { 8 => 15, -5 => 19, 4 => 10, 5 => 5, -7 => 1, 93 => 27, _ => 0 };
        internal static nuint CatalogColumnSize(short sqlType) => sqlType == -9 ? 32 * 1024 * 1024 : ColumnSize(sqlType);
        internal static short Precision(short sqlType) => sqlType switch { 8 => 53, -5 => 19, 4 => 10, 5 => 5, -7 => 1, 93 => 7, _ => 0 };
        internal static nint OctetLength(short sqlType) => sqlType switch { 0 or -9 or -3 => -4, -7 => 1, 5 => 2, 4 => 4, 93 => 16, _ => 8 };
        internal static nint CatalogOctetLength(short sqlType) => sqlType == -9 ? 64 * 1024 * 1024 : OctetLength(sqlType);
        internal static nint DisplaySize(short sqlType) => sqlType switch { 8 => 24, -5 => 20, 4 => 11, 5 => 6, 0 or -9 or -3 => -4, _ => (nint)ColumnSize(sqlType) };
        internal static string TypeName(short sqlType) => sqlType switch { 0 => "", 8 => "DOUBLE", -5 => "BIGINT", 4 => "INTEGER", 5 => "SMALLINT", -7 => "BIT", 93 => "TIMESTAMP", -3 => "VARBINARY", _ => "WVARCHAR" };

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
                    if (targetType == -2 && value is not byte[])
                        return statement.Fail("07006", "Binary output requires a binary SQL value.");
                    string text = targetType == -2 ? "" : Text(value);
                    byte[] bytes = targetType == -2 ? (byte[])value : NativeText.Bytes(text, wide);
                    int offset = advance ? statement.DataOffset : 0;
                    if (offset >= bytes.Length && offset > 0) return 100;
                    int remaining = bytes.Length - offset;
                    if (indicator != 0) *(nint*)indicator = remaining;
                    if (buffer == 0 || capacity == 0) {
                        if (remaining == 0) return 0;
                        statement.Diagnostics.Add(new("01004", "Column data was truncated; no value bytes were copied."));
                        return 1;
                    }
                    int terminator = targetType == -2 ? 0 : wide ? 2 : 1;
                    if (capacity < terminator) return statement.Fail("HY090", "Buffer cannot hold a string terminator.");
                    int copied = (int)Math.Min(remaining, Math.Min(int.MaxValue, capacity - terminator));
                    if (wide) copied -= copied % 2;
                    if (targetType != -2 && value is byte[]) copied -= copied % (wide ? 4 : 2);
                    if (targetType != -2 && value is not string && SqlType(reader, column - 1) is 8 or -5 or 4 or 5 or -7) {
                        int point = text.IndexOf('.');
                        int whole = text.Contains('E') || point < 0 ? text.Length : point;
                        if (offset == 0 && copied < whole * (wide ? 2 : 1))
                            return statement.Fail("22003", "Character buffer cannot preserve the whole numeric value.");
                    }
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
                        if (value is string bitText) {
                            decimal wholeBit = TruncateText(bitText, out fractional);
                            if (wholeBit < 0 || wholeBit >= 2 || fractional && bitText.TrimStart().StartsWith('-')) throw new OverflowException();
                            *(byte*)buffer = (byte)wholeBit; size = 1; break;
                        }
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
                        double floating = Floating(value);
                        float single = (float)floating;
                        if (!float.IsFinite(single) || single == 0 && floating != 0) throw new OverflowException();
                        *(float*)buffer = single; size = 4; break;
                    case 8: *(double*)buffer = Floating(value); size = 8; break;
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
            if (value is string text) return TruncateText(text, out fractional);
            object result = value switch {
                double number => Math.Truncate(number), float number => MathF.Truncate(number),
                decimal number => decimal.Truncate(number), _ => value
            };
            fractional = !result.Equals(value);
            return result;
        }

        private static decimal TruncateText(string value, out bool fractional) {
            System.Text.RegularExpressions.Match match = NumericLiteral(value);
            string digits = match.Groups["whole"].Value + match.Groups["fraction"].Value;
            int first = digits.AsSpan().IndexOfAnyExcept('0');
            if (first < 0) { fractional = false; return 0; }
            string exponentText = match.Groups["exponent"].Value;
            int exponent = 0;
            if (exponentText.Length != 0 && !int.TryParse(exponentText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out exponent))
                exponent = exponentText[0] == '-' ? int.MinValue : int.MaxValue;
            long count = (long)match.Groups["whole"].Length + exponent - first;
            if (count > 20) throw new OverflowException();
            digits = digits[first..];
            int whole = (int)Math.Max(0, count);
            fractional = whole < digits.Length && digits.AsSpan(whole).IndexOfAnyExcept('0') >= 0;
            // At most 20 whole digits can fit any supported integer target. Truncate the
            // literal before parsing so long fractions never round an exact boundary.
            string integral = whole == 0 ? "0" : whole <= digits.Length ? digits[..whole] : digits.PadRight(whole, '0');
            return decimal.Parse(match.Groups["sign"].Value + integral, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        }

        private static System.Text.RegularExpressions.Match NumericLiteral(string value) {
            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(value.Trim(),
                @"\A(?<sign>[+-]?)(?<whole>[0-9]*)(?:\.(?<fraction>[0-9]*))?(?:[eE](?<exponent>[+-]?[0-9]+))?\z",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant | System.Text.RegularExpressions.RegexOptions.NonBacktracking,
                TimeSpan.FromSeconds(1));
            string digits = match.Groups["whole"].Value + match.Groups["fraction"].Value;
            if (!match.Success || digits.Length == 0) throw new FormatException("Expected a numeric literal.");
            return match;
        }

        private static double Floating(object value) {
            System.Text.RegularExpressions.Match? literal = value is string text ? NumericLiteral(text) : null;
            double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (!double.IsFinite(number)) throw new OverflowException();
            if (number == 0 && literal is not null &&
                (literal.Groups["whole"].Value + literal.Groups["fraction"].Value).AsSpan().IndexOfAnyExcept('0') >= 0)
                throw new OverflowException();
            return number;
        }

        internal static object? Read(ParameterBinding binding) {
            nint length = binding.Indicator != 0 ? *(nint*)binding.Indicator : binding.Type is 1 or -8 ? -3 : binding.Length;
            if (length == -1) return DBNull.Value;
            if (length < 0 && length != -3) throw new NotSupportedException("Data-at-execution parameters are not supported.");
            if (binding.Buffer == 0) throw new ArgumentException("Parameter value buffer is required.");
            return NativeParameter.CheckValue(binding, binding.Type switch {
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
            });
        }

        private static DateTime ReadDate(nint buffer, bool timestamp) {
            ushort* parts = (ushort*)buffer;
            DateTime result;
            try {
                result = new DateTime(*(short*)buffer, parts[1], parts[2], timestamp ? parts[3] : 0, timestamp ? parts[4] : 0, timestamp ? parts[5] : 0);
            } catch (ArgumentOutOfRangeException) {
                throw new ParameterValueException("22008", "Date or timestamp fields are outside the supported calendar range.");
            }
            if (!timestamp) return result;
            uint fraction = *(uint*)((byte*)buffer + 12);
            if (fraction >= 1_000_000_000 || fraction % 100 != 0)
                throw new ParameterValueException("22008", "Timestamp fraction must be representable in 100 nanosecond ticks.");
            return result.AddTicks(fraction / 100);
        }

        private static string Text(object value) => value switch {
            byte[] bytes => Convert.ToHexString(bytes),
            bool bit => bit ? "1" : "0",
            DateTime date => date.ToString("O", CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
        };
    }
}
