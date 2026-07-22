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
        private bool TryInsertSimpleObjectRowsAsDeferredDirectSave<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            Type rowType = typeof(T);
            if (rowType == typeof(object) || rowType.IsInterface || rowType.IsAbstract) {
                return false;
            }

            SimpleObjectExportPlan plan = GetSimpleObjectExportPlan(rowType);
            if (!plan.CanUseDirectSave) {
                return false;
            }

            string[] headers = plan.Headers;
            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Length)) {
                return false;
            }

            SimpleObjectExportValueGetter[] getters = plan.Getters;
            var values = new object?[checked(rows.Count * getters.Length)];
            for (int r = 0; r < rows.Count; r++) {
                object? row = rows[r];
                if (row == null || row.GetType() != rowType) {
                    return false;
                }

                int rowOffset = r * getters.Length;
                for (int c = 0; c < getters.Length; c++) {
                    values[rowOffset + c] = getters[c](row);
                }
            }

            string range = BuildObjectExportRange(startRow, headers.Length, rows.Count, includeHeaders);
            return TryInsertCellValuesAsDeferredDirectSave(Name, headers, plan.ColumnTypes, values, headers.Length, rows.Count, startRow, includeHeaders, range);
        }

        private bool TryInsertSimpleObjectRowsAsCellValues<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            Type rowType = typeof(T);
            if (rowType == typeof(object) || rowType.IsInterface || rowType.IsAbstract) {
                return false;
            }

            SimpleObjectExportPlan plan = GetSimpleObjectExportPlan(rowType);
            if (!plan.CanUseDirectSave) {
                return false;
            }

            string[] headers = plan.Headers;
            SimpleObjectExportValueGetter[] getters = plan.Getters;
            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * headers.Length);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, headers[c]);
                }

                row++;
            }

            for (int r = 0; r < rows.Count; r++) {
                object? item = rows[r];
                if (item == null || item.GetType() != rowType) {
                    return false;
                }

                for (int c = 0; c < getters.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, getters[c](item) ?? string.Empty);
                }

                row++;
            }

            CellValues(cells);
            return true;
        }

        private static SimpleObjectExportPlan GetSimpleObjectExportPlan(
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type) {
            if (SimpleObjectExportPlans.TryGetValue(type, out SimpleObjectExportPlan? existing)) {
                return existing;
            }

            SimpleObjectExportPlan created = CreateSimpleObjectExportPlan(type);
            return SimpleObjectExportPlans.GetOrAdd(type, created);
        }

        private static SimpleObjectExportPlan CreateSimpleObjectExportPlan(
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type) {
            var properties = GetSimpleObjectExportProperties(type);
            if (properties.Length == 0) {
                return SimpleObjectExportPlan.NotSupported;
            }

            var headers = new string[properties.Length];
            var getters = new SimpleObjectExportValueGetter[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
                headers[i] = properties[i].Name;
                getters[i] = CreateSimpleObjectExportValueGetter(properties[i]);
            }

            if (HasDuplicateObjectExportHeaders(headers)) {
                return SimpleObjectExportPlan.NotSupported;
            }

            return new SimpleObjectExportPlan(headers, getters, InferSimpleObjectExportColumnTypes(properties), canUseDirectSave: true);
        }

        private static SimpleObjectExportValueGetter CreateSimpleObjectExportValueGetter(PropertyInfo property)
            => row => property.GetValue(row, null);

        private static PropertyInfo[] GetSimpleObjectExportProperties(
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type) {
            var properties = type.GetProperties().Where(property => property.CanRead).ToArray();
            if (properties.Length == 0) {
                return Array.Empty<PropertyInfo>();
            }

            for (int i = 0; i < properties.Length; i++) {
                if (properties[i].GetIndexParameters().Length != 0
                    || !IsObjectExportScalarType(properties[i].PropertyType)) {
                    return Array.Empty<PropertyInfo>();
                }
            }

            return properties;
        }

        private static Type[] InferSimpleObjectExportColumnTypes(IReadOnlyList<PropertyInfo> properties) {
            var columnTypes = new Type[properties.Count];
            for (int i = 0; i < columnTypes.Length; i++) {
                columnTypes[i] = Nullable.GetUnderlyingType(properties[i].PropertyType) ?? properties[i].PropertyType;
            }

            return columnTypes;
        }

        private static bool IsObjectExportScalarType(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid)
#if NET6_0_OR_GREATER
                || type == typeof(DateOnly)
                || type == typeof(TimeOnly)
#endif
                ;
        }

        private sealed class SimpleObjectExportPlan {
            internal static readonly SimpleObjectExportPlan NotSupported = new(
                Array.Empty<string>(),
                Array.Empty<SimpleObjectExportValueGetter>(),
                Array.Empty<Type>(),
                canUseDirectSave: false);

            internal SimpleObjectExportPlan(
                string[] headers,
                SimpleObjectExportValueGetter[] getters,
                Type[] columnTypes,
                bool canUseDirectSave) {
                Headers = headers;
                Getters = getters;
                ColumnTypes = columnTypes;
                CanUseDirectSave = canUseDirectSave;
            }

            internal string[] Headers { get; }

            internal SimpleObjectExportValueGetter[] Getters { get; }

            internal Type[] ColumnTypes { get; }

            internal bool CanUseDirectSave { get; }
        }

        private delegate object? SimpleObjectExportValueGetter(object? row);
    }
}
