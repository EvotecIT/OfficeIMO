using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader
    {
        /// <summary>
        /// Reads a rectangular range and maps rows (excluding the header row) into instances of T.
        /// Header cells are matched to public writable properties on T by name (case-insensitive).
        /// </summary>
        public IEnumerable<T> ReadObjects<T>(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) where T : new()
        {
            var values = ReadRange(a1Range, mode, ct);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows == 0 || cols == 0) yield break;

            // Build property map from headers
            var headers = new string[cols];
            for (int c = 0; c < cols; c++)
                headers[c] = values[0, c]?.ToString() ?? $"Column{c + 1}";

            var props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanWrite)
                .ToDictionary(p => p.Name, p => p, StringComparer.OrdinalIgnoreCase);

            var map = new (int ColIndex, PropertyInfo Prop)[cols];
            for (int c = 0; c < cols; c++)
            {
                if (_opt.NormalizeHeaders)
                    headers[c] = System.Text.RegularExpressions.Regex.Replace(headers[c], "\\s+", " ").Trim();
                if (props.TryGetValue(headers[c], out var pi))
                    map[c] = (c, pi);
                else
                    map[c] = (-1, null!);
            }

            for (int r = 1; r < rows; r++)
            {
                if (ct.IsCancellationRequested) yield break;
                var obj = new T();
                for (int c = 0; c < cols; c++)
                {
                    if (map[c].ColIndex == -1) continue;
                    var pi = map[c].Prop;
                    var raw = values[r, c];
                    if (raw is null) continue;

                    object? converted = TryChangeType(raw, pi.PropertyType, _opt.Culture);
                    if (converted is not null || IsNullableType(pi.PropertyType))
                        pi.SetValue(obj, converted);
                }
                yield return obj;
            }
        }

        private static bool IsNullableType(Type t)
        {
            return !t.IsValueType || Nullable.GetUnderlyingType(t) != null;
        }

        private object? TryChangeType(object value, Type targetType, CultureInfo culture)
        {
            if (value == null) return null;
            var srcType = value.GetType();
            if (targetType.IsAssignableFrom(srcType)) return value;

            var nullable = Nullable.GetUnderlyingType(targetType);
            var destType = nullable ?? targetType;

            // Custom type-converter hook
            var hook = _opt.TypeConverter;
            if (hook != null)
            {
                var (ok, v) = hook(value, destType, culture);
                if (ok) return v;
            }

            try
            {
                if (destType == typeof(string)) return Convert.ToString(value, culture);
                if (destType == typeof(bool)) return Convert.ToBoolean(value, culture);
                if (destType == typeof(int)) return Convert.ToInt32(value, culture);
                if (destType == typeof(long)) return Convert.ToInt64(value, culture);
                if (destType == typeof(double)) return Convert.ToDouble(value, culture);
                if (destType == typeof(decimal)) return Convert.ToDecimal(value, culture);
                if (destType == typeof(DateTime))
                {
                    if (value is DateTime dt) return dt;
                    if (value is double oa) return DateTime.FromOADate(oa);
                    if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                    return null;
                }
                // Fallback to ChangeType
                return Convert.ChangeType(value, destType, culture);
            }
            catch
            {
                return null;
            }
        }
    }
}
