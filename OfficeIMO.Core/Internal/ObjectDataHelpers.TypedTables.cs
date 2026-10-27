using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using OfficeIMO.Data;

namespace OfficeIMO.Core.Internal;

internal static partial class ObjectDataHelpers {
    /// <summary>Projects a declared row schema without reflecting over arbitrary runtime object graphs.</summary>
    internal static TypedTableData ProjectTypedTable<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> data, bool includeHeaders, int maximumRows = int.MaxValue,
        int maximumColumns = int.MaxValue, long maximumCells = long.MaxValue, OfficeTableOptions? options = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        ValidateTableAlignments(options);
        using var enumerator = data.GetEnumerator();
        bool hasRow = enumerator.MoveNext();
        if (!hasRow && ObjectDictionaryAdapter.IsDictionaryType(typeof(T)))
            throw new ArgumentException("Empty dictionary rows have no column schema. Supply explicit column selectors.", nameof(data));
        List<ObjectDictionaryEntry> entries = new();
        bool dictionary = hasRow && ObjectDictionaryAdapter.TryGetEntries(enumerator.Current, int.MaxValue, out entries);
        bool firstDictionaryRow = dictionary;
        MemberInfo[] members = dictionary ? Array.Empty<MemberInfo>() : GetTypedTableMembers(typeof(T));
        string[] names = dictionary ? entries.Select(entry => entry.Key as string
                ?? throw new ArgumentException("Dictionary table keys must be strings.", nameof(data))).ToArray()
            : members.Select(member => member.Name).ToArray();
        names = SelectTableColumns(names, options);
        if (names.Length == 0 || names.Length > maximumColumns)
            throw new ArgumentException("The row schema has no columns or exceeds the native column limit.", nameof(data));
        if (names.Any(string.IsNullOrWhiteSpace) || names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Length)
            throw new ArgumentException("Table column names must be non-empty and unique.", nameof(data));
        string[] headers = names.Select(name => GetTableHeader(name, options)).ToArray();
        if (!dictionary) members = names.Select(name => members.First(member => member.Name == name)).ToArray();
        var rows = new List<object?[]>();
        while (hasRow) {
            if (rows.Count >= maximumRows || ((long)rows.Count + 1 + (includeHeaders ? 1 : 0)) * headers.Length > maximumCells)
                throw new InvalidOperationException("The table exceeds the native row or cell limit.");
            T row = enumerator.Current;
            if (row is null) throw new ArgumentException("Table rows must not contain null entries.", nameof(data));
            Dictionary<string, object?>? dictionaryValues = null;
            if (dictionary) {
                if (!firstDictionaryRow && !ObjectDictionaryAdapter.TryGetEntries(row, int.MaxValue, out entries))
                    throw new ArgumentException("All dictionary rows must use the same table schema.", nameof(data));
                firstDictionaryRow = false;
                dictionaryValues = new Dictionary<string, object?>(StringComparer.Ordinal);
                foreach (var entry in entries) {
                    if (entry.Key is not string key) throw new ArgumentException("Dictionary table keys must be strings.", nameof(data));
                    dictionaryValues[key] = entry.Value;
                }
            }
            var values = new object?[headers.Length];
            for (int column = 0; column < headers.Length; column++) {
                if (dictionaryValues != null) {
                    dictionaryValues.TryGetValue(names[column], out values[column]);
                } else {
                    values[column] = members[column] is PropertyInfo property
                        ? property.GetValue(row, null) : ((FieldInfo)members[column]).GetValue(row);
                }
                values[column] = FormatTableValue(names[column], values[column], options);
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
