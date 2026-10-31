using System.Data;
using System.Diagnostics.CodeAnalysis;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an editable native table from typed properties, fields, or string-keyed dictionary columns.</summary>
    public ExcelTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null, int row = 1, int column = 1) {
        ValidateTypedTableOrigin(row, column);
        ObjectDataHelpers.ValidateTableAlignments(options);
        bool headers = options?.IncludeHeaders ?? true;
        var excelOptions = options as ExcelTableOptions;
        var projection = new ExcelObjectTableProjection(this, row, column, headers);
        string? range = projection.WriteRows(rows, excelOptions?.Projection, options);
        if (range == null) {
            if (!headers) throw new ArgumentException("A headerless table requires at least one row.", nameof(rows));
            var schema = ObjectDataHelpers.ProjectTypedTable(Array.Empty<T>(), true,
                maximumColumns: ExcelA1.MaxColumns - column + 1, options: options);
            if (excelOptions?.Projection is { } limits) {
                ExcelObjectTableProjection.EnsureRowsFromColumnLimit(schema.Headers.Length, limits.MaxColumns);
                if (schema.Headers.Length > limits.MaxCells)
                    throw new InvalidDataException($"The table header exceeds the {limits.MaxCells}-cell materialization limit.");
            }
            for (int field = 0; field < schema.Headers.Length; field++) CellValue(row, column + field, schema.Headers[field]);
            range = ExcelA1.CellReference(row, column) + ":" + ExcelA1.CellReference(row, column + schema.Headers.Length - 1);
        }
        var table = new ExcelTable(this, AddTableAndGetName(range, headers, options?.Name ?? "Table",
            excelOptions?.Style ?? ExcelTableStyle.TableStyleMedium2,
            excelOptions?.IncludeAutoFilter ?? true, headerNames: GetTypedTableColumnNames(projection.ColumnNames, headers, excelOptions)));
        FinalizeTypedTable(table, options);
        return table;
    }

    /// <summary>Adds an editable native table from a data table, preserving its typed cell values.</summary>
    public ExcelTable AddTable(DataTable data, OfficeTableOptions? options = null, int row = 1, int column = 1) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        ValidateTypedTableOrigin(row, column);
        ObjectDataHelpers.ValidateTableAlignments(options);
        bool headers = options?.IncludeHeaders ?? true;
        int count = checked(data.Rows.Count + (headers ? 1 : 0));
        if (count == 0 || data.Columns.Count == 0 || count > ExcelA1.MaxRows - row + 1)
            throw new ArgumentException("The data table is empty or exceeds the native worksheet bounds.", nameof(data));
        using DataTable? projected = ExcelObjectTableProjection.HasTableProjection(options)
            ? ProjectDataTable(data, options!, ExcelA1.MaxColumns - column + 1) : null;
        data = projected ?? data;
        if (data.Columns.Count > ExcelA1.MaxColumns - column + 1)
            throw new ArgumentException("The data table is empty or exceeds the native worksheet bounds.", nameof(data));
        InsertDataTable(data, row, column, headers, (options as ExcelTableOptions)?.ExecutionMode,
            (options as ExcelTableOptions)?.CancellationToken ?? default);
        string range = ExcelA1.CellReference(row, column) + ":" + ExcelA1.CellReference(row + count - 1, column + data.Columns.Count - 1);
        var table = new ExcelTable(this, AddTableAndGetName(range, headers, options?.Name ?? data.TableName,
            (options as ExcelTableOptions)?.Style ?? ExcelTableStyle.TableStyleMedium2,
            (options as ExcelTableOptions)?.IncludeAutoFilter ?? true,
            headerNames: GetTypedTableColumnNames(data.Columns.Cast<DataColumn>().Select(field => field.ColumnName).ToArray(),
                headers, options as ExcelTableOptions)));
        FinalizeTypedTable(table, options);
        return table;
    }

    private static void ValidateTypedTableOrigin(int row, int column) {
        if (row < 1 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
    }
}
