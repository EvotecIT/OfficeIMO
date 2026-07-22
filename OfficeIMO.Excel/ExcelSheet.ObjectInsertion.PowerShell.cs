using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryInsertPowerShellObjectRowsAsCellValues<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0 || rows[0] == null) {
                return false;
            }

            object first = rows[0]!;
            Type rowType = first.GetType();
            if (!IsPowerShellObjectExportType(rowType)) {
                return false;
            }

            PowerShellObjectExportPlan plan = PowerShellObjectExportPlans.GetOrAdd(rowType, CreatePowerShellObjectExportPlan);
            if (!plan.CanProject) {
                return false;
            }

            var propertyPlanCache = new PowerShellPropertyExportPlanCache();
            var headers = new List<string>();
            var state = new FlatDictionaryProjectionState {
                InferredColumnTypes = new Type?[BufferedPowerShellObjectExportInitialColumnCapacity]
            };
            object?[] firstRow = new object?[BufferedPowerShellObjectExportInitialColumnCapacity];
            if (!TryProjectPowerShellObjectFirstRow(first, plan, ref propertyPlanCache, headers, state, ref firstRow)) {
                return false;
            }

            state.NormalizeColumnTypeWidth(headers.Count);
            if (headers.Count == 0
                || headers.Count > BufferedPowerShellObjectExportColumnLimit
                || includeHeaders && state.HasBlankDisplayHeader
                || HasDuplicateObjectExportHeaders(headers)
                || !CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                return false;
            }

            int columnCount = headers.Count;
            string[] headerArray = headers.ToArray();
            bool rowTypeIsSealed = rowType.IsSealed;
            var values = new object?[checked(rows.Count * columnCount)];
            for (int c = 0; c < columnCount; c++) {
                values[c] = c < firstRow.Length ? firstRow[c] : null;
            }

            for (int r = 1; r < rows.Count; r++) {
                object? row = rows[r];
                if (row == null || (!rowTypeIsSealed && row.GetType() != rowType)) {
                    return false;
                }

                int rowOffset = r * columnCount;
                if (!TryProjectPowerShellObjectExistingRow(
                    row,
                    plan,
                    ref propertyPlanCache,
                    headerArray,
                    state.InferredColumnTypes,
                    values,
                    rowOffset,
                    columnCount)) {
                    return false;
                }
            }

            string range = BuildObjectExportRange(startRow, columnCount, rows.Count, includeHeaders);
            return TryInsertCellValuesAsDeferredDirectSave(
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                values,
                columnCount,
                rows.Count,
                startRow,
                includeHeaders,
                range);

        }

        private static bool TryProjectPowerShellObjectFirstRow(
            object item,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            List<string> headers,
            FlatDictionaryProjectionState state,
            ref object?[] firstRow) {
            if (!TryGetPowerShellObjectProperties(item, plan, out object? propertiesValue)) {
                return false;
            }

            bool added = false;
            if (propertiesValue is object?[] propertyArray) {
                for (int i = 0; i < propertyArray.Length; i++) {
                    if (!TryProjectPowerShellObjectFirstProperty(propertyArray[i], plan, ref propertyPlanCache, headers, state, ref firstRow, ref added)) {
                        return false;
                    }
                }

                return added;
            }

            if (propertiesValue is System.Collections.IList propertyList) {
                for (int i = 0; i < propertyList.Count; i++) {
                    if (!TryProjectPowerShellObjectFirstProperty(propertyList[i], plan, ref propertyPlanCache, headers, state, ref firstRow, ref added)) {
                        return false;
                    }
                }

                return added;
            }

            if (propertiesValue is not IEnumerable properties) {
                return false;
            }

            foreach (object? property in properties) {
                if (!TryProjectPowerShellObjectFirstProperty(property, plan, ref propertyPlanCache, headers, state, ref firstRow, ref added)) {
                    return false;
                }
            }

            return added;
        }

        private static bool TryProjectPowerShellObjectExistingRow(
            object item,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            string[] headers,
            Type?[] inferredColumnTypes,
            object?[] values,
            int rowOffset,
            int columnCount) {
            if (!TryGetPowerShellObjectProperties(item, plan, out object? propertiesValue)) {
                return false;
            }

            int propertyIndex = 0;
            if (propertiesValue is object?[] propertyArray) {
                for (int i = 0; i < propertyArray.Length; i++) {
                    if (!TryProjectPowerShellObjectExistingProperty(propertyArray[i], plan, ref propertyPlanCache, headers, inferredColumnTypes, values, rowOffset, columnCount, ref propertyIndex)) {
                        return false;
                    }
                }

                return propertyIndex > 0;
            }

            if (propertiesValue is System.Collections.IList propertyList) {
                for (int i = 0; i < propertyList.Count; i++) {
                    if (!TryProjectPowerShellObjectExistingProperty(propertyList[i], plan, ref propertyPlanCache, headers, inferredColumnTypes, values, rowOffset, columnCount, ref propertyIndex)) {
                        return false;
                    }
                }

                return propertyIndex > 0;
            }

            if (propertiesValue is not IEnumerable properties) {
                return false;
            }

            foreach (object? property in properties) {
                if (!TryProjectPowerShellObjectExistingProperty(property, plan, ref propertyPlanCache, headers, inferredColumnTypes, values, rowOffset, columnCount, ref propertyIndex)) {
                    return false;
                }
            }

            return propertyIndex > 0;
        }

        private static bool TryGetPowerShellObjectProperties(
            object item,
            PowerShellObjectExportPlan plan,
            out object? propertiesValue) {
            if (!plan.CanProject) {
                propertiesValue = null;
                return false;
            }

            try {
                propertiesValue = plan.PropertiesGetter(item);
                return true;
            } catch {
                propertiesValue = null;
                return false;
            }
        }

        private static bool TryProjectPowerShellObjectFirstProperty(
            object? property,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            List<string> headers,
            FlatDictionaryProjectionState state,
            ref object?[] firstRow,
            ref bool added) {
            if (property == null) {
                return true;
            }

            Type propertyType = property.GetType();
            PowerShellPropertyExportPlan propertyPlan = propertyPlanCache.Get(plan, propertyType);
            if (!propertyPlan.CanProject) {
                return true;
            }

            try {
                if (propertyPlan.IsGettableGetter != null && !propertyPlan.IsGettableGetter(property)) {
                    return true;
                }

                string? name = propertyPlan.NameGetter(property);
                object? value = propertyPlan.ValueGetter(property);
                if (!TryGetFlatObjectExportValueType(value, out Type? valueType)) {
                    return false;
                }

                string columnName = name ?? string.Empty;
                if (string.IsNullOrWhiteSpace(columnName)) {
                    state.HasBlankDisplayHeader = true;
                }

                int columnIndex = headers.IndexOf(columnName);
                if (columnIndex < 0) {
                    if (headers.Count >= BufferedPowerShellObjectExportColumnLimit) {
                        return false;
                    }

                    columnIndex = headers.Count;
                    EnsurePowerShellObjectFirstRowCapacity(ref firstRow, state, columnIndex + 1);
                    headers.Add(columnName);
                }

                firstRow[columnIndex] = value;
                UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, valueType);
                added = true;
                return true;
            } catch {
                return true;
            }
        }

        private static void EnsurePowerShellObjectFirstRowCapacity(
            ref object?[] firstRow,
            FlatDictionaryProjectionState state,
            int requiredLength) {
            if (firstRow.Length >= requiredLength) {
                return;
            }

            int newLength = firstRow.Length * 2;
            if (newLength < requiredLength) {
                newLength = requiredLength;
            }

            if (newLength > BufferedPowerShellObjectExportColumnLimit) {
                newLength = BufferedPowerShellObjectExportColumnLimit;
            }

            Array.Resize(ref firstRow, newLength);
            state.EnsureColumnTypeCapacity(newLength);
        }

        private static bool TryProjectPowerShellObjectExistingProperty(
            object? property,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            string[] headers,
            Type?[] inferredColumnTypes,
            object?[] values,
            int rowOffset,
            int columnCount,
            ref int propertyIndex) {
            if (property == null) {
                return true;
            }

            Type propertyType = property.GetType();
            PowerShellPropertyExportPlan propertyPlan = propertyPlanCache.Get(plan, propertyType);
            if (!propertyPlan.CanProject) {
                return true;
            }

            try {
                if (propertyPlan.IsGettableGetter != null && !propertyPlan.IsGettableGetter(property)) {
                    return true;
                }

                string? name = propertyPlan.NameGetter(property);
                object? value = propertyPlan.ValueGetter(property);
                if (!TryGetFlatObjectExportValueType(value, out Type? valueType)) {
                    return false;
                }

                string columnName = name ?? string.Empty;
                int columnIndex;
                if ((uint)propertyIndex < (uint)columnCount
                    && string.Equals(columnName, headers[propertyIndex], StringComparison.Ordinal)) {
                    columnIndex = propertyIndex;
                } else {
                    columnIndex = Array.IndexOf(headers, columnName);
                }

                if (columnIndex < 0) {
                    return false;
                }

                values[rowOffset + columnIndex] = value;
                UpdateObjectExportColumnType(inferredColumnTypes, columnIndex, valueType);
                propertyIndex++;
                return true;
            } catch {
                return true;
            }
        }


        private static bool TryProjectPowerShellObjectRow(object item, TryAddObjectExportValue tryAddValue) {
            Type itemType = item.GetType();
            if (!IsPowerShellObjectExportType(itemType)) {
                return false;
            }

            PowerShellObjectExportPlan plan = PowerShellObjectExportPlans.GetOrAdd(itemType, CreatePowerShellObjectExportPlan);
            var propertyPlanCache = new PowerShellPropertyExportPlanCache();
            return TryProjectPowerShellObjectRow(item, plan, ref propertyPlanCache, tryAddValue);
        }

        private static bool TryProjectPowerShellObjectRow(
            object item,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            TryAddObjectExportValue tryAddValue) {
            return TryProjectPowerShellObjectRow(
                item,
                plan,
                ref propertyPlanCache,
                (propertyIndex, name, value) => tryAddValue(name, value));
        }

        private static bool TryProjectPowerShellObjectRow(
            object item,
            PowerShellObjectExportPlan plan,
            ref PowerShellPropertyExportPlanCache propertyPlanCache,
            TryAddIndexedObjectExportValue tryAddValue) {
            if (!plan.CanProject) {
                return false;
            }

            object? propertiesValue;
            try {
                propertiesValue = plan.PropertiesGetter(item);
            } catch {
                return false;
            }

            var localPropertyPlanCache = propertyPlanCache;
            try {
                bool added = false;
                int propertyIndex = 0;

                if (propertiesValue is object?[] propertyArray) {
                    for (int i = 0; i < propertyArray.Length; i++) {
                        if (!TryProjectProperty(propertyArray[i])) {
                            return false;
                        }
                    }

                    return added;
                }

                if (propertiesValue is System.Collections.IList propertyList) {
                    for (int i = 0; i < propertyList.Count; i++) {
                        if (!TryProjectProperty(propertyList[i])) {
                            return false;
                        }
                    }

                    return added;
                }

                if (propertiesValue is not IEnumerable properties) {
                    return false;
                }

                foreach (object? property in properties) {
                    if (!TryProjectProperty(property)) {
                        return false;
                    }
                }

                return added;

                bool TryProjectProperty(object? property) {
                    if (property == null) {
                        return true;
                    }

                    Type propertyType = property.GetType();
                    PowerShellPropertyExportPlan propertyPlan = localPropertyPlanCache.Get(plan, propertyType);
                    if (!propertyPlan.CanProject) {
                        return true;
                    }

                    try {
                        if (propertyPlan.IsGettableGetter != null && !propertyPlan.IsGettableGetter(property)) {
                            return true;
                        }

                        string? name = propertyPlan.NameGetter(property);
                        object? value = propertyPlan.ValueGetter(property);
                        if (!tryAddValue(propertyIndex, name, value)) {
                            return false;
                        }

                        added = true;
                        propertyIndex++;
                    } catch {
                    }

                    return true;
                }
            } finally {
                propertyPlanCache = localPropertyPlanCache;
            }
        }

        private struct PowerShellPropertyExportPlanCache {
            private Type? _lastPropertyType;
            private PowerShellPropertyExportPlan? _lastPropertyPlan;

            internal PowerShellPropertyExportPlan Get(PowerShellObjectExportPlan plan, Type propertyType) {
                if (_lastPropertyType == propertyType && _lastPropertyPlan != null) {
                    return _lastPropertyPlan;
                }

                _lastPropertyType = propertyType;
                _lastPropertyPlan = plan.GetPropertyPlan(propertyType);
                return _lastPropertyPlan;
            }
        }

        private static bool IsPowerShellObjectExportType(Type type)
            => string.Equals(type.FullName, "System.Management.Automation.PSObject", StringComparison.Ordinal)
               || string.Equals(type.FullName, "System.Management.Automation.PSCustomObject", StringComparison.Ordinal);

        [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "This branch is reached only for PowerShell's runtime object model through InsertObjects, which is explicitly marked as reflection-based compatibility API.")]
        private static PowerShellObjectExportPlan CreatePowerShellObjectExportPlan(Type type) {
            PropertyInfo? properties = type.GetProperty("Properties", BindingFlags.Public | BindingFlags.Instance);
            if (properties == null || !typeof(IEnumerable).IsAssignableFrom(properties.PropertyType)) {
                return PowerShellObjectExportPlan.NotSupported;
            }

            return new PowerShellObjectExportPlan(CreatePowerShellValueGetter(properties));
        }

        [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "This branch is reached only for PowerShell's runtime property model through InsertObjects, which is explicitly marked as reflection-based compatibility API.")]
        private static PowerShellPropertyExportPlan CreatePowerShellPropertyExportPlan(Type type) {
            PropertyInfo? name = type.GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo? value = type.GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo? isGettable = type.GetProperty("IsGettable", BindingFlags.Public | BindingFlags.Instance);
            if (name == null || value == null) {
                return PowerShellPropertyExportPlan.NotSupported;
            }

            return new PowerShellPropertyExportPlan(
                CreatePowerShellStringGetter(name),
                CreatePowerShellValueGetter(value),
                isGettable == null ? null : CreatePowerShellBooleanGetter(isGettable));
        }

        private static Func<object, object?> CreatePowerShellValueGetter(PropertyInfo property)
            => row => property.GetValue(row, null);

        private static Func<object, string?> CreatePowerShellStringGetter(PropertyInfo property)
            => row => property.GetValue(row, null)?.ToString();

        private static Func<object, bool> CreatePowerShellBooleanGetter(PropertyInfo property)
            => row => property.GetValue(row, null) is bool value && value;

        private sealed class PowerShellObjectExportPlan {
            internal static readonly PowerShellObjectExportPlan NotSupported = new();

            private readonly ConcurrentDictionary<Type, PowerShellPropertyExportPlan> _propertyPlans = new();

            private PowerShellObjectExportPlan() {
                PropertiesGetter = _ => null;
                CanProject = false;
            }

            internal PowerShellObjectExportPlan(Func<object, object?> propertiesGetter) {
                PropertiesGetter = propertiesGetter;
                CanProject = true;
            }

            internal bool CanProject { get; }

            internal Func<object, object?> PropertiesGetter { get; }

            internal PowerShellPropertyExportPlan GetPropertyPlan(Type propertyType)
                => _propertyPlans.GetOrAdd(propertyType, CreatePowerShellPropertyExportPlan);
        }

        private sealed class PowerShellPropertyExportPlan {
            internal static readonly PowerShellPropertyExportPlan NotSupported = new();

            private PowerShellPropertyExportPlan() {
                NameGetter = _ => null;
                ValueGetter = _ => null;
                CanProject = false;
            }

            internal PowerShellPropertyExportPlan(
                Func<object, string?> nameGetter,
                Func<object, object?> valueGetter,
                Func<object, bool>? isGettableGetter) {
                NameGetter = nameGetter;
                ValueGetter = valueGetter;
                IsGettableGetter = isGettableGetter;
                CanProject = true;
            }

            internal bool CanProject { get; }

            internal Func<object, string?> NameGetter { get; }

            internal Func<object, object?> ValueGetter { get; }

            internal Func<object, bool>? IsGettableGetter { get; }
        }
    }
}
