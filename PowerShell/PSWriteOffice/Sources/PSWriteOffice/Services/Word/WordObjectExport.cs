using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Globalization;
using System.Linq;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Services.Word;

/// <summary>Projects export rows through the shared tabular input service and .NET display formatting.</summary>
internal static class WordObjectExport {
    internal static IReadOnlyList<object?> ProjectRows(IReadOnlyCollection<object?> input, string[]? properties, Hashtable? formats) {
        using DataTable data = ExcelTabularInputService.ToDataTable(TableInputCollector.RequireRows(input, "InputObject"));
        if (data.Rows.Count == 0 || data.Columns.Count == 0) {
            throw new ArgumentException("Provide at least one data row with a property.", nameof(input));
        }

        DataColumn[] columns = properties == null
            ? data.Columns.Cast<DataColumn>().ToArray()
            : properties.Select(property => data.Columns[property] ?? throw new ArgumentException("Unknown property: " + property, "Property")).ToArray();
        if (columns.Length == 0 || columns.Distinct().Count() != columns.Length) {
            throw new ArgumentException("Choose at least one distinct table property.", "Property");
        }
        var columnFormats = new Dictionary<DataColumn, string>();
        if (formats != null) {
            foreach (DictionaryEntry entry in formats) {
                string name = Convert.ToString(entry.Key, CultureInfo.InvariantCulture) ?? string.Empty;
                DataColumn? column = columns.FirstOrDefault(item => item.ColumnName.Equals(name, StringComparison.OrdinalIgnoreCase));
                string format = Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                if (column == null || string.IsNullOrWhiteSpace(format)) {
                    throw new ArgumentException("ColumnFormat requires a selected column name and a nonempty .NET format string: " + name, "ColumnFormat");
                }
                columnFormats.Add(column, format);
            }
        }

        var rows = new List<object?>(data.Rows.Count);
        foreach (DataRow row in data.Rows) {
            var values = new OrderedDictionary(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in columns) {
                object? value = row.IsNull(column) ? null : row[column];
                if (value is IFormattable formattable && columnFormats.TryGetValue(column, out string? format)) {
                    try {
                        value = formattable.ToString(format, CultureInfo.InvariantCulture);
                    } catch (FormatException error) {
                        throw new ArgumentException("Invalid .NET format for column " + column.ColumnName + ": " + format, "ColumnFormat", error);
                    }
                }
                values.Add(column.ColumnName, value);
            }
            rows.Add(values);
        }
        return rows;
    }
}
