using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Core.Internal;

internal static partial class ObjectDataHelpers {
    /// <summary>Projects a declared row schema without reflecting over arbitrary runtime object graphs.</summary>
    internal static TypedTableData ProjectTypedTable<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> data, bool includeHeaders, int maximumRows = int.MaxValue,
        int maximumColumns = int.MaxValue, long maximumCells = long.MaxValue) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        using var enumerator = data.GetEnumerator();
        bool hasRow = enumerator.MoveNext();
        IDictionary? dictionary = hasRow ? enumerator.Current as IDictionary : null;
        MemberInfo[] members = dictionary == null ? GetTypedTableMembers(typeof(T)) : Array.Empty<MemberInfo>();
        string[] headers = dictionary == null ? members.Select(member => member.Name).ToArray()
            : dictionary.Keys.Cast<object>().Select(key => key as string
                ?? throw new ArgumentException("Dictionary table keys must be strings.", nameof(data))).ToArray();
        if (headers.Length == 0 || headers.Length > maximumColumns)
            throw new ArgumentException("The row schema has no columns or exceeds the native column limit.", nameof(data));
        if (headers.Any(string.IsNullOrWhiteSpace) || headers.Distinct(StringComparer.OrdinalIgnoreCase).Count() != headers.Length)
            throw new ArgumentException("Table column names must be non-empty and unique.", nameof(data));
        var rows = new List<object?[]>();
        while (hasRow) {
            if (rows.Count >= maximumRows || ((long)rows.Count + 1 + (includeHeaders ? 1 : 0)) * headers.Length > maximumCells)
                throw new InvalidOperationException("The table exceeds the native row or cell limit.");
            T row = enumerator.Current;
            if (row is null) throw new ArgumentException("Table rows must not contain null entries.", nameof(data));
            var values = new object?[headers.Length];
            for (int column = 0; column < headers.Length; column++) {
                if (dictionary != null) {
                    if (row is not IDictionary rowDictionary)
                        throw new ArgumentException("All dictionary rows must use the same table schema.", nameof(data));
                    values[column] = rowDictionary.Contains(headers[column]) ? rowDictionary[headers[column]] : null;
                } else {
                    values[column] = members[column] is PropertyInfo property
                        ? property.GetValue(row, null) : ((FieldInfo)members[column]).GetValue(row);
                }
            }
            rows.Add(values);
            hasRow = enumerator.MoveNext();
        }
        return new TypedTableData(headers, rows);
    }

    private static MemberInfo[] GetTypedTableMembers(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] Type type) =>
        type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.GetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
            .Cast<MemberInfo>().Concat(type.GetFields(BindingFlags.Instance | BindingFlags.Public)).ToArray();

    internal sealed class TypedTableData(string[] headers, IReadOnlyList<object?[]> rows) {
        internal string[] Headers { get; } = headers;
        internal IReadOnlyList<object?[]> Rows { get; } = rows;
    }
}
