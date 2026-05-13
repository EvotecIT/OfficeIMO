using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

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
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (rows <= 1 || cols == 0) return Array.Empty<T>();

            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = rows * cols;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                decided = policy.Decide("ReadObjectsAs", workload);
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel) {
                if (TryReadObjectsSequentialSinglePass<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var fastResult)) {
                    return fastResult;
                }

                return ReadObjectsSequential<T>(a1Range, r1, c1, r2, c2, rows, cols, ct);
            }

            var rawCells = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadObjectsAs", decided, ct, workload);

            // Build property map from normalized, disambiguated headers so repeated
            // source headers remain addressable instead of colliding.
            var headerValues = new object?[cols];
            foreach (var cell in rawCells) {
                if (cell.Row != r1) {
                    continue;
                }

                int cc = cell.Col - c1;
                if ((uint)cc < (uint)cols) {
                    headerValues[cc] = cell.TypedValue;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);

            var headerBindings = GetTypedHeaderBindings<T>(headers, a1Range);
            var map = headerBindings.Bindings;
            bool canCancel = ct.CanBeCanceled;

            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            int dataRowCount = rows - 1;
            var result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                result.Add(new T());
            }

            bool hasCustomConverters = _opt.CellValueConverter != null || _opt.TypeConverter != null;
            for (int i = 0; i < rawCells.Count; i++) {
                if (canCancel && (i & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                var cell = rawCells[i];
                if (cell.Row <= r1 || cell.TypedValue is null) {
                    continue;
                }

                int rr = cell.Row - r1 - 1;
                int cc = cell.Col - c1;
                if ((uint)rr >= (uint)result.Count || (uint)cc >= (uint)cols) {
                    continue;
                }

                var binding = map[cc];
                if (binding == null) {
                    continue;
                }

                object? converted = TryChangeType(cell.TypedValue, binding, _opt.Culture);
                if (converted is null
                    && !hasCustomConverters
                    && ShouldRetryRawDateStyledNumericBinding(cell, binding)
                    && TryConvertRawForBinding(cell, binding, out object? rawConverted)) {
                    converted = rawConverted;
                }

                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (converted is not null || binding.IsNullable) {
                    binding.SetValue(result[rr], converted);
                }
            }

            return result;
        }

        /// <summary>
        /// Streams a rectangular range and maps each data row into an instance of T without materializing the full result set first.
        /// Header cells are matched to public writable properties on T by name (case-insensitive).
        /// Enumerate the returned sequence while the owning reader is still open.
        /// </summary>
        public IEnumerable<T> ReadObjectsStream<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(string a1Range, CancellationToken ct = default) where T : new() {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (rows <= 1 || cols == 0) {
                return Array.Empty<T>();
            }

            return ReadObjectsStreamIterator<T>(a1Range, r1, c1, r2, c2, cols, ct);
        }

        private IEnumerable<T> ReadObjectsStreamIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            Dictionary<int, Row>? pendingRows = null;
            int nextDataRow = r1 + 1;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (bindings == null) {
                    if (rowIndex == r1) {
                        bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);

                        while (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                            pendingRows.Remove(nextDataRow);
                            var pendingTarget = new T();
                            FillTypedObjectFromRow(pendingRow, c1, c2, bindings, pendingTarget, ct, ref convertedCells);
                            yield return pendingTarget;
                            nextDataRow++;
                        }

                        continue;
                    }

                    pendingRows ??= new Dictionary<int, Row>();
                    pendingRows[rowIndex] = row;
                    continue;
                }

                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex < nextDataRow) {
                    continue;
                }

                if (rowIndex > nextDataRow) {
                    pendingRows ??= new Dictionary<int, Row>();
                    pendingRows[rowIndex] = row;
                    continue;
                }

                var currentRow = row;
                while (true) {
                    var target = new T();
                    FillTypedObjectFromRow(currentRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                    nextDataRow++;

                    if (pendingRows == null || !pendingRows.TryGetValue(nextDataRow, out currentRow)) {
                        break;
                    }

                    pendingRows.Remove(nextDataRow);
                }
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            while (nextDataRow <= r2) {
                if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                    pendingRows.Remove(nextDataRow);
                    var target = new T();
                    FillTypedObjectFromRow(pendingRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                } else {
                    yield return new T();
                }

                nextDataRow++;
            }
        }

        private bool TryReadObjectsSequentialSinglePass<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<T> result) where T : new() {
            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                result.Add(new T());
            }

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (rowIndex == r1) {
                    bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                    continue;
                }

                if (bindings == null) {
                    return false;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return bindings != null;
        }

        private IEnumerable<T> ReadObjectsSequential<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex != r1) {
                    continue;
                }

                bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                break;
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);

            int dataRowCount = rows - 1;
            var result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                result.Add(new T());
            }

            int convertedCells = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return result;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            string a1Range,
            int c1,
            int c2,
            int cols) where T : new() {
            var headerValues = new object?[cols];
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                if (TryConvertCell(cell, out object? value)) {
                    headerValues[columnIndex - c1] = value;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromMissingRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int cols) where T : new() {
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, static _ => null, _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private void FillTypedObjectFromRow<T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            int c1,
            int c2,
            TypedPropertyBinding<T>?[] bindings,
            T target,
            CancellationToken ct,
            ref int convertedCells) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                if (canCancel && (++convertedCells & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                var binding = bindings[columnIndex - c1];
                if (binding == null) {
                    continue;
                }

                bool convertedSuccessfully = TryConvertCellForBinding(cell, binding, out object? converted);
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (convertedSuccessfully) {
                    binding.SetValue(target, converted);
                }
            }
        }

        private TypedHeaderBindingCache<T> GetTypedHeaderBindings<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string[] headers,
            string a1Range) where T : new() {
            var propertyMaps = TypedObjectBindingCache<T>.PropertyMaps;
            string typeName = propertyMaps.TypeName;
            foreach (string diagnostic in propertyMaps.Diagnostics) {
                _opt.Execution.ReportInfo(diagnostic);
            }

            var headerBindings = TypedObjectBindingCache<T>.GetHeaderBindings(headers);

            if (_opt.StrictTypedMapping) {
                var strictIssues = new List<string>(propertyMaps.Diagnostics);
                strictIssues.AddRange(headerBindings.UnmappedIssues);
                if (strictIssues.Count > 0) {
                    throw new InvalidOperationException(
                        $"Typed mapping for '{typeName}' is strict and could not resolve all headers in range '{a1Range}'. " +
                        string.Join(" ", strictIssues));
                }
            }

            return headerBindings;
        }

        private sealed class TypedPropertyBinding<TTarget> {
            internal TypedPropertyBinding(
                PropertyInfo property,
                Type propertyType,
                Type destinationType,
                bool isNullable,
                bool needsDateStyleConversion,
                Action<TTarget, object?> setValue,
                Func<object, CultureInfo, object?> convertValue) {
                Property = property;
                PropertyType = propertyType;
                DestinationType = destinationType;
                IsNullable = isNullable;
                NeedsDateStyleConversion = needsDateStyleConversion;
                SetValue = setValue;
                ConvertValue = convertValue;
            }

            internal PropertyInfo Property { get; }
            internal Type PropertyType { get; }
            internal Type DestinationType { get; }
            internal bool IsNullable { get; }
            internal bool NeedsDateStyleConversion { get; }
            internal Action<TTarget, object?> SetValue { get; }
            internal Func<object, CultureInfo, object?> ConvertValue { get; }
        }

        private sealed class TypedPropertyMapCache {
            internal TypedPropertyMapCache(
                string typeName,
                Dictionary<string, PropertyInfo> exactProperties,
                Dictionary<string, PropertyInfo> exactAliases,
                Dictionary<string, PropertyInfo> canonicalProperties,
                Dictionary<string, PropertyInfo> canonicalAliases,
                IReadOnlyList<string> diagnostics) {
                TypeName = typeName;
                ExactProperties = exactProperties;
                ExactAliases = exactAliases;
                CanonicalProperties = canonicalProperties;
                CanonicalAliases = canonicalAliases;
                Diagnostics = diagnostics;
            }

            internal string TypeName { get; }
            internal Dictionary<string, PropertyInfo> ExactProperties { get; }
            internal Dictionary<string, PropertyInfo> ExactAliases { get; }
            internal Dictionary<string, PropertyInfo> CanonicalProperties { get; }
            internal Dictionary<string, PropertyInfo> CanonicalAliases { get; }
            internal IReadOnlyList<string> Diagnostics { get; }
        }

        private sealed class TypedHeaderBindingCache<TTarget> where TTarget : new() {
            internal TypedHeaderBindingCache(TypedPropertyBinding<TTarget>?[] bindings, IReadOnlyList<string> unmappedIssues) {
                Bindings = bindings;
                UnmappedIssues = unmappedIssues;
            }

            internal TypedPropertyBinding<TTarget>?[] Bindings { get; }
            internal IReadOnlyList<string> UnmappedIssues { get; }
        }

        private static class TypedObjectBindingCache<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] TTarget> where TTarget : new() {
            private const int HeaderBindingCacheLimit = 64;

            internal static readonly PropertyInfo[] WritableProperties = typeof(TTarget)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanWrite)
                .ToArray();

            internal static readonly Dictionary<PropertyInfo, TypedPropertyBinding<TTarget>> Bindings = WritableProperties
                .ToDictionary(prop => prop, CreateBinding);

            internal static readonly TypedPropertyMapCache PropertyMaps = CreatePropertyMaps();

            private static readonly ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>> HeaderBindings = new ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>>(StringComparer.Ordinal);

            internal static TypedHeaderBindingCache<TTarget> GetHeaderBindings(string[] headers) {
                string key = CreateHeaderBindingKey(headers);
                if (HeaderBindings.TryGetValue(key, out var cached)) {
                    return cached;
                }

                var created = CreateHeaderBindings(headers);
                if (HeaderBindings.Count < HeaderBindingCacheLimit) {
                    return HeaderBindings.GetOrAdd(key, created);
                }

                return created;
            }

            private static TypedPropertyBinding<TTarget> CreateBinding(PropertyInfo property) {
                var nullable = Nullable.GetUnderlyingType(property.PropertyType);
                var destinationType = nullable ?? property.PropertyType;
                return new TypedPropertyBinding<TTarget>(
                    property,
                    property.PropertyType,
                    destinationType,
                    !property.PropertyType.IsValueType || nullable != null,
                    NeedsDateStyleConversion(destinationType),
                    CreateSetter(property),
                    CreateConverter(destinationType));
            }

            private static bool NeedsDateStyleConversion(Type destinationType) {
                return destinationType == typeof(DateTime) || destinationType == typeof(string);
            }

            private static Func<object, CultureInfo, object?> CreateConverter(Type destinationType) {
                if (destinationType == typeof(string)) {
                    return static (value, culture) => value as string ?? Convert.ToString(value, culture);
                }

                if (destinationType == typeof(bool)) {
                    return static (value, culture) => {
                        try {
                            if (value is bool boolValue) return boolValue;
                            return Convert.ToBoolean(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(int)) {
                    return static (value, culture) => {
                        try {
                            if (value is int intValue) return intValue;
                            if (value is double doubleValue
                                && doubleValue >= int.MinValue
                                && doubleValue <= int.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (int)doubleValue;
                            }

                            return Convert.ToInt32(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(long)) {
                    return static (value, culture) => {
                        try {
                            if (value is long longValue) return longValue;
                            if (value is double doubleValue
                                && doubleValue >= long.MinValue
                                && doubleValue <= long.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (long)doubleValue;
                            }

                            return Convert.ToInt64(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(double)) {
                    return static (value, culture) => {
                        try {
                            if (value is double doubleValue) return doubleValue;
                            return Convert.ToDouble(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(decimal)) {
                    return static (value, culture) => {
                        try {
                            if (value is decimal decimalValue) return decimalValue;
                            return Convert.ToDecimal(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(DateTime)) {
                    return static (value, culture) => {
                        try {
                            if (value is DateTime dt) return dt;
                            if (value is double oa) return DateTime.FromOADate(oa);
                            if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                            return null;
                        } catch {
                            return null;
                        }
                    };
                }

                return (value, culture) => ConvertToDestinationType(value, destinationType, culture);
            }

            private static Action<TTarget, object?> CreateSetter(PropertyInfo property) {
                try {
                    var target = Expression.Parameter(typeof(TTarget), "target");
                    var value = Expression.Parameter(typeof(object), "value");
                    var converted = Expression.Convert(value, property.PropertyType);
                    var body = Expression.Assign(Expression.Property(target, property), converted);
                    return Expression.Lambda<Action<TTarget, object?>>(body, target, value).Compile();
                } catch {
                    return (target, value) => property.SetValue(target, value);
                }
            }

            private static TypedPropertyMapCache CreatePropertyMaps() {
                string typeName = typeof(TTarget).Name;
                var diagnostics = new List<string>();

                return new TypedPropertyMapCache(
                    typeName,
                    BuildPropertyMap(WritableProperties, prop => new[] { prop.Name }, diagnostics, typeName, "exact property"),
                    BuildPropertyMap(WritableProperties, GetPropertyAliases, diagnostics, typeName, "explicit alias"),
                    BuildPropertyMap(WritableProperties, prop => new[] { CanonicalizeMemberName(prop.Name) }, diagnostics, typeName, "friendly property"),
                    BuildPropertyMap(WritableProperties, prop => GetPropertyAliases(prop).Select(CanonicalizeMemberName), diagnostics, typeName, "friendly alias"),
                    diagnostics.ToArray());
            }

            private static TypedHeaderBindingCache<TTarget> CreateHeaderBindings(string[] headers) {
                var map = new TypedPropertyBinding<TTarget>?[headers.Length];
                var assignedProps = new HashSet<PropertyInfo>();

                // Exact property matches win first so alias/friendly fallback does not steal
                // a property from a later exact-name column.
                for (int c = 0; c < headers.Length; c++) {
                    if (PropertyMaps.ExactProperties.TryGetValue(headers[c], out var pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                // Explicit aliases come next (DisplayName/DataMember/ExcelColumn).
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    if (PropertyMaps.ExactAliases.TryGetValue(headers[c], out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string canonicalHeader = CanonicalizeMemberName(headers[c]);
                    if (canonicalHeader.Length == 0) {
                        continue;
                    }

                    if (PropertyMaps.CanonicalProperties.TryGetValue(canonicalHeader, out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                        continue;
                    }

                    if (PropertyMaps.CanonicalAliases.TryGetValue(canonicalHeader, out pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                var unmappedIssues = new List<string>();
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string header = headers[c];
                    if (header.Length == 0) {
                        continue;
                    }

                    unmappedIssues.Add($"[TypedRead UnmappedHeader] Type='{PropertyMaps.TypeName}', header='{header}', column={c + 1}.");
                }

                return new TypedHeaderBindingCache<TTarget>(map, unmappedIssues.ToArray());
            }

            private static string CreateHeaderBindingKey(string[] headers) {
                var builder = new StringBuilder(headers.Length * 16);
                for (int i = 0; i < headers.Length; i++) {
                    string header = headers[i] ?? string.Empty;
                    builder.Append(header.Length.ToString(CultureInfo.InvariantCulture));
                    builder.Append(':');
                    builder.Append(header);
                    builder.Append('|');
                }

                return builder.ToString();
            }
        }

        private object? TryChangeType<TTarget>(object value, TypedPropertyBinding<TTarget> binding, CultureInfo culture) {
            if (value == null) return null;
            var srcType = value.GetType();
            if (binding.PropertyType.IsAssignableFrom(srcType)) return value;

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, binding.DestinationType, culture);
                if (ok) return v;
            }

            return binding.ConvertValue(value, culture);
        }

        private bool TryConvertCellForBinding<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (_opt.CellValueConverter != null || _opt.TypeConverter != null) {
                object? value = ConvertCell(cell);
                if (value is null) {
                    return binding.IsNullable;
                }

                converted = TryChangeType(value, binding, _opt.Culture);
                return converted is not null || binding.IsNullable;
            }

            CellValues? typeHint = cell.DataType?.Value;
            bool hasFormula = cell.CellFormula is not null;
            string? formulaText = hasFormula ? ExtractFormulaText(cell) : null;
            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            string? rawText = preferFormulaText ? null : ExtractRawText(cell);
            string? inlineText = preferFormulaText ? null : ExtractInlineString(cell, typeHint);

            if (rawText == null && inlineText == null && formulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (hasFormula && (!_opt.UseCachedFormulaResult || rawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = formulaText ?? rawText ?? inlineText;
                    return converted is not null || binding.IsNullable;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (!string.IsNullOrEmpty(inlineText)
                && ReturnBindingConversion(TryConvertStringForBinding(inlineText, binding, out converted), binding, converted)) {
                return true;
            }

            if (typeHint == CellValues.SharedString) {
                string? text = rawText;
                if (TryParseSharedStringIndex(rawText, out int sstIndex)) {
                    text = _sst.Get(sstIndex);
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(text, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Boolean && rawText != null) {
                if (ReturnBindingConversion(TryConvertBooleanForBinding(rawText == "1", binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.String || typeHint == CellValues.InlineString) {
                if (ReturnBindingConversion(TryConvertStringForBinding(rawText ?? inlineText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Date && rawText != null) {
                if (DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)
                    && ReturnBindingConversion(TryConvertDateTimeForBinding(dt, binding, out converted), binding, converted)) {
                    return true;
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (rawText == null) {
                return binding.IsNullable;
            }

            uint? styleIndex = null;
            if (_opt.TreatDatesUsingNumberFormat && binding.NeedsDateStyleConversion) {
                styleIndex = cell.StyleIndex?.Value;
                if (styleIndex is not null && _styles.IsDateLike(styleIndex.Value)) {
                    if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa)
                        && ReturnBindingConversion(TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted), binding, converted)) {
                        return true;
                    }

                    if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                        return true;
                    }

                    return TryConvertCellForBindingFallback(cell, binding, out converted);
                }
            }

            if (ReturnBindingConversion(TryConvertNumericTextForBinding(rawText, binding, out converted), binding, converted)) {
                return true;
            }

            return TryConvertCellForBindingFallback(cell, binding, out converted);
        }

        private bool TryConvertCellForBindingFallback<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            var raw = SnapshotCell(cell);
            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (TryConvertRawForBinding(raw, binding, out converted)) {
                return converted is not null || binding.IsNullable;
            }

            object? typedValue = ConvertRaw(raw).TypedValue;
            if (typedValue is null) {
                return binding.IsNullable;
            }

            converted = TryChangeType(typedValue, binding, _opt.Culture);
            return converted is not null || binding.IsNullable;
        }

        private static bool ReturnBindingConversion<TTarget>(
            bool convertedByFastPath,
            TypedPropertyBinding<TTarget> binding,
            object? converted) {
            return convertedByFastPath && (converted is not null || binding.IsNullable);
        }

        private bool TryConvertRawForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (raw.HasFormula && (!_opt.UseCachedFormulaResult || raw.RawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                    return true;
                }

                return false;
            }

            if (!string.IsNullOrEmpty(raw.InlineText)) {
                return TryConvertStringForBinding(raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                if (!TryParseSharedStringIndex(raw.RawText, out int sstIndex)) {
                    return TryConvertStringForBinding(raw.RawText, binding, out converted);
                }

                return TryConvertStringForBinding(_sst.Get(sstIndex), binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean && raw.RawText != null) {
                return TryConvertBooleanForBinding(raw.RawText == "1", binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return TryConvertStringForBinding(raw.RawText ?? raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date && raw.RawText != null) {
                if (DateTime.TryParse(raw.RawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                    return TryConvertDateTimeForBinding(dt, binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            if (raw.RawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && raw.StyleIndex is not null
                && _styles.IsDateLike(raw.StyleIndex.Value)) {
                if (double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa)) {
                    return TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            return TryConvertNumericTextForBinding(raw.RawText, binding, out converted);
        }

        private bool ShouldRetryRawDateStyledNumericBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding) {
            if (!_opt.TreatDatesUsingNumberFormat
                || binding.NeedsDateStyleConversion
                || !IsNumericBindingDestination(binding.DestinationType)
                || raw.RawText == null
                || raw.StyleIndex is null
                || !_styles.IsDateLike(raw.StyleIndex.Value)) {
                return false;
            }

            return double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _);
        }

        private static bool IsNumericBindingDestination(Type destinationType) {
            return destinationType == typeof(int)
                || destinationType == typeof(long)
                || destinationType == typeof(double)
                || destinationType == typeof(decimal);
        }

        private bool TryConvertNumericTextForBinding<TTarget>(
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;

            if (destinationType == typeof(int)) {
                if (int.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out int intValue)) {
                    converted = intValue;
                    return true;
                }

                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                    && doubleValue >= int.MinValue
                    && doubleValue <= int.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (int)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(long)) {
                if (long.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out long longValue)) {
                    converted = longValue;
                    return true;
                }

                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                    && doubleValue >= long.MinValue
                    && doubleValue <= long.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (long)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(double)) {
                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)) {
                    converted = doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(decimal)) {
                if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out decimal decimalValue)) {
                    converted = decimalValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(bool)) {
                if (rawText == "1") {
                    converted = true;
                    return true;
                }

                if (rawText == "0") {
                    converted = false;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(string)) {
                converted = rawText;
                return true;
            }

            return false;
        }

        private bool TryConvertStringForBinding<TTarget>(
            string? text,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            if (text == null) {
                return binding.IsNullable;
            }

            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(string)) {
                converted = text;
                return true;
            }

            if (destinationType == typeof(bool) && bool.TryParse(text, out bool boolValue)) {
                converted = boolValue;
                return true;
            }

            if (destinationType == typeof(DateTime)
                && DateTime.TryParse(text, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                converted = dt;
                return true;
            }

            return TryConvertNumericTextForBinding(text, binding, out converted);
        }

        private static bool TryConvertBooleanForBinding<TTarget>(
            bool value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(bool)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString();
                return true;
            }

            return false;
        }

        private bool TryConvertDateTimeForBinding<TTarget>(
            DateTime value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(DateTime)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString(_opt.Culture);
                return true;
            }

            return false;
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

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, destType, culture);
                if (ok) return v;
            }

            return ConvertToDestinationType(value, destType, culture);
        }

        private static object? ConvertToDestinationType(object value, Type destType, CultureInfo culture) {
            try {
                if (destType == typeof(string)) {
                    return value as string ?? Convert.ToString(value, culture);
                }

                if (destType == typeof(bool)) {
                    if (value is bool boolValue) return boolValue;
                    return Convert.ToBoolean(value, culture);
                }

                if (destType == typeof(int)) {
                    if (value is int intValue) return intValue;
                    if (value is double doubleValue
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (int)doubleValue;
                    }

                    return Convert.ToInt32(value, culture);
                }

                if (destType == typeof(long)) {
                    if (value is long longValue) return longValue;
                    if (value is double doubleValue
                        && doubleValue >= long.MinValue
                        && doubleValue <= long.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (long)doubleValue;
                    }

                    return Convert.ToInt64(value, culture);
                }

                if (destType == typeof(double)) {
                    if (value is double doubleValue) return doubleValue;
                    return Convert.ToDouble(value, culture);
                }

                if (destType == typeof(decimal)) {
                    if (value is decimal decimalValue) return decimalValue;
                    return Convert.ToDecimal(value, culture);
                }

                if (destType == typeof(DateTime)) {
                    if (value is DateTime dt) return dt;
                    if (value is double oa) return DateTime.FromOADate(oa);
                    if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                    return null;
                }

                return Convert.ChangeType(value, destType, culture);
            } catch {
                return null;
            }
        }
    }
}
