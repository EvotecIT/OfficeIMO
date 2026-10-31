using System.Data;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an attached native table using explicit AOT-safe column selectors and the ordinary table options.</summary>
    /// <remarks>Table column names must be nonblank and unique. Use WriteData when duplicate or blank display headers are needed.</remarks>
    public ExcelTable AddTable<T>(IEnumerable<T> rows,
        (string Header, Func<T, object?> Selector)[] columns,
        OfficeTableOptions? options = null, int row = 1, int column = 1) {
        if (rows == null) throw new ArgumentNullException(nameof(rows));
        if (columns == null || columns.Length == 0) throw new ArgumentException("At least one column selector is required.", nameof(columns));
        ValidateTypedTableOrigin(row, column);
        ObjectDataHelpers.ValidateTableAlignments(options);
        string[] sourceNames = columns.Select((field, index) => field.Header ?? $"Column{index + 1}").ToArray();
        string[] selectedNames = ObjectDataHelpers.SelectTableColumns(sourceNames, options);
        var selected = selectedNames.Select(name => columns[Array.IndexOf(sourceNames, name)]).ToArray();
        if (selected.Length == 0 || selected.Length > ExcelA1.MaxColumns - column + 1)
            throw new ArgumentException("The selected columns are empty or exceed the native worksheet bounds.", nameof(columns));
        using var data = new DataTable();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < selected.Length; i++) {
            string header = selectedNames[i];
            if (string.IsNullOrWhiteSpace(header) || !names.Add(header))
                throw new ArgumentException("Table column names must be nonblank and unique.", nameof(columns));
            data.Columns.Add(header, typeof(object));
        }
        bool includeHeaders = options?.IncludeHeaders ?? true;
        var cancellation = (options as ExcelTableOptions)?.CancellationToken ?? default;
        cancellation.ThrowIfCancellationRequested();
        IReadOnlyList<T> items = MaterializeObjectRowsBounded(rows, includeHeaders, row, cancellation);
        ValidateObjectExportDenseDimensions(items.Count, selected.Length, includeHeaders);
        foreach (T item in items) {
            cancellation.ThrowIfCancellationRequested();
            var values = new object?[selected.Length];
            for (int i = 0; i < selected.Length; i++) {
                cancellation.ThrowIfCancellationRequested();
                values[i] = selected[i].Selector?.Invoke(item) ?? DBNull.Value;
            }
            data.Rows.Add(values);
        }
        // Keep original member keys so the DataTable owner applies headers and formatters exactly once.
        cancellation.ThrowIfCancellationRequested();
        return AddTable(data, options, row, column);
    }
}
