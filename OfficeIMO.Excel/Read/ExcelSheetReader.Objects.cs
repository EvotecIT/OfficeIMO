using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a rectangular range and maps rows (excluding the header row) into instances of T.
        /// Header cells are matched to public writable properties on T by name (case-insensitive).
        /// </summary>
        public IEnumerable<T> ReadObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) where T : new() {
            var values = ReadRange(a1Range, mode, ct);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows == 0 || cols == 0) yield break;

            // Build property map from normalized, disambiguated headers so repeated
            // source headers remain addressable instead of colliding.
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), _opt.NormalizeHeaders);

            var writableProps = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanWrite)
                .ToArray();

            string typeName = typeof(T).Name;
            var mappingDiagnostics = new List<string>();

            var exactProps = BuildPropertyMap(writableProps, prop => new[] { prop.Name }, mappingDiagnostics, typeName, "exact property");
            var exactAliases = BuildPropertyMap(writableProps, GetPropertyAliases, mappingDiagnostics, typeName, "explicit alias");
            var canonicalProps = BuildPropertyMap(writableProps, prop => new[] { CanonicalizeMemberName(prop.Name) }, mappingDiagnostics, typeName, "friendly property");
            var canonicalAliases = BuildPropertyMap(writableProps, prop => GetPropertyAliases(prop).Select(CanonicalizeMemberName), mappingDiagnostics, typeName, "friendly alias");

            foreach (string diagnostic in mappingDiagnostics) {
                _opt.Execution.ReportInfo(diagnostic);
            }

            var map = new (int ColIndex, PropertyInfo Prop)[cols];
            var assignedProps = new HashSet<PropertyInfo>();

            // Exact property matches win first so alias/friendly fallback does not steal
            // a property from a later exact-name column.
            for (int c = 0; c < cols; c++) {
                if (exactProps.TryGetValue(headers[c], out var pi)) {
                    map[c] = (c, pi);
                    assignedProps.Add(pi);
                } else {
                    map[c] = (-1, null!);
                }
            }

            // Explicit aliases come next (DisplayName/DataMember/ExcelColumn).
            for (int c = 0; c < cols; c++) {
                if (map[c].ColIndex != -1) {
                    continue;
                }

                if (exactAliases.TryGetValue(headers[c], out var pi) && !assignedProps.Contains(pi)) {
                    map[c] = (c, pi);
                    assignedProps.Add(pi);
                }
            }

            for (int c = 0; c < cols; c++) {
                if (map[c].ColIndex != -1) {
                    continue;
                }

                string canonicalHeader = CanonicalizeMemberName(headers[c]);
                if (canonicalHeader.Length == 0) {
                    continue;
                }

                if (canonicalProps.TryGetValue(canonicalHeader, out var pi) && !assignedProps.Contains(pi)) {
                    map[c] = (c, pi);
                    assignedProps.Add(pi);
                    continue;
                }

                if (canonicalAliases.TryGetValue(canonicalHeader, out pi) && !assignedProps.Contains(pi)) {
                    map[c] = (c, pi);
                    assignedProps.Add(pi);
                }
            }

            if (_opt.StrictTypedMapping) {
                var strictIssues = new List<string>(mappingDiagnostics);
                for (int c = 0; c < cols; c++) {
                    if (map[c].ColIndex != -1) {
                        continue;
                    }

                    string header = headers[c];
                    if (header.Length == 0) {
                        continue;
                    }

                    strictIssues.Add($"[TypedRead UnmappedHeader] Type='{typeName}', header='{header}', column={c + 1}.");
                }

                if (strictIssues.Count > 0) {
                    throw new InvalidOperationException(
                        $"Typed mapping for '{typeName}' is strict and could not resolve all headers in range '{a1Range}'. " +
                        string.Join(" ", strictIssues));
                }
            }

            for (int r = 1; r < rows; r++) {
                if (ct.IsCancellationRequested) yield break;
                var obj = new T();
                for (int c = 0; c < cols; c++) {
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

        private static bool IsNullableType(Type t) {
            return !t.IsValueType || Nullable.GetUnderlyingType(t) != null;
        }

        private static Dictionary<string, PropertyInfo> BuildPropertyMap(
            IEnumerable<PropertyInfo> props,
            Func<PropertyInfo, IEnumerable<string>> candidateFactory,
            ICollection<string> diagnostics,
            string typeName,
            string mappingKind) {
            var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            var ambiguous = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ambiguousProps = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var prop in props) {
                foreach (string rawCandidate in candidateFactory(prop)) {
                    if (string.IsNullOrWhiteSpace(rawCandidate)) {
                        continue;
                    }

                    string candidate = rawCandidate;
                    if (candidate.Length == 0 || ambiguous.Contains(candidate)) {
                        continue;
                    }

                    if (map.TryGetValue(candidate, out var existing) && existing != prop) {
                        map.Remove(candidate);
                        ambiguous.Add(candidate);
                        if (!ambiguousProps.TryGetValue(candidate, out var propNames)) {
                            propNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { existing.Name };
                            ambiguousProps[candidate] = propNames;
                        }

                        propNames.Add(prop.Name);
                        continue;
                    }

                    if (ambiguousProps.TryGetValue(candidate, out var existingAmbiguousNames)) {
                        existingAmbiguousNames.Add(prop.Name);
                        continue;
                    }

                    map[candidate] = prop;
                }
            }

            foreach (var pair in ambiguousProps.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                diagnostics.Add(
                    $"[TypedRead AmbiguousMapping] Type='{typeName}', match='{mappingKind}', header='{pair.Key}', properties='{string.Join(", ", pair.Value.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}'.");
            }

            return map;
        }

        private static IEnumerable<string> GetPropertyAliases(PropertyInfo propertyInfo) {
            var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void YieldIfUnique(string? candidate, List<string> buffer) {
                if (!string.IsNullOrWhiteSpace(candidate)) {
                    string text = candidate!;
                    if (yielded.Add(text)) {
                        buffer.Add(text);
                    }
                }
            }

            var aliases = new List<string>();

            var displayName = propertyInfo.GetCustomAttribute<DisplayNameAttribute>(inherit: true);
            YieldIfUnique(displayName?.DisplayName, aliases);

            var dataMember = propertyInfo.GetCustomAttribute<DataMemberAttribute>(inherit: true);
            YieldIfUnique(dataMember?.Name, aliases);

            var excelColumn = propertyInfo.GetCustomAttribute<ExcelColumnAttribute>(inherit: true);
            if (excelColumn != null) {
                YieldIfUnique(excelColumn.Name, aliases);
                foreach (string alias in excelColumn.Aliases) {
                    YieldIfUnique(alias, aliases);
                }
            }

            return aliases;
        }

        private static string CanonicalizeMemberName(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            string text = value ?? string.Empty;
            var builder = new StringBuilder(text.Length);
            foreach (char character in text) {
                if (char.IsLetterOrDigit(character)) {
                    builder.Append(char.ToUpperInvariant(character));
                }
            }

            return builder.ToString();
        }

        private object? TryChangeType(object value, Type targetType, CultureInfo culture) {
            if (value == null) return null;
            var srcType = value.GetType();
            if (targetType.IsAssignableFrom(srcType)) return value;

            var nullable = Nullable.GetUnderlyingType(targetType);
            var destType = nullable ?? targetType;

            // Custom type-converter hook
            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, destType, culture);
                if (ok) return v;
            }

            try {
                if (destType == typeof(string)) return Convert.ToString(value, culture);
                if (destType == typeof(bool)) return Convert.ToBoolean(value, culture);
                if (destType == typeof(int)) return Convert.ToInt32(value, culture);
                if (destType == typeof(long)) return Convert.ToInt64(value, culture);
                if (destType == typeof(double)) return Convert.ToDouble(value, culture);
                if (destType == typeof(decimal)) return Convert.ToDecimal(value, culture);
                if (destType == typeof(DateTime)) {
                    if (value is DateTime dt) return dt;
                    if (value is double oa) return DateTime.FromOADate(oa);
                    if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                    return null;
                }
                // Fallback to ChangeType
                return Convert.ChangeType(value, destType, culture);
            } catch {
                return null;
            }
        }
    }
}
