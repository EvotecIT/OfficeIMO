using System.Data;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    private void ApplyTableAlignments(ExcelTable table, OfficeTableOptions? options) {
        if (options == null || options.Alignments.Count == 0 || table.Range == null) return;
        var (firstRow, firstColumn, lastRow, lastColumn) = ExcelA1.ParseRange(table.Range);
        for (int column = firstColumn; column <= lastColumn && column - firstColumn < options.Alignments.Count; column++) {
            if (options.Alignments[column - firstColumn] is not { } alignment) continue;
            var native = alignment switch {
                OfficeAlignment.Left or OfficeAlignment.Start => ExcelHorizontalAlignment.Left,
                OfficeAlignment.Center => ExcelHorizontalAlignment.Center,
                OfficeAlignment.Right or OfficeAlignment.End => ExcelHorizontalAlignment.Right,
                _ => ExcelHorizontalAlignment.Justify
            };
            for (int row = firstRow; row <= lastRow; row++) CellAlign(row, column, native);
        }
    }

    private static DataTable ProjectDataTable(DataTable data, OfficeTableOptions options, int maximumColumns) {
        string[] names = ObjectDataHelpers.SelectTableColumns(data.Columns.Cast<DataColumn>().Select(column => column.ColumnName), options);
        if (names.Length == 0 || names.Length > maximumColumns)
            throw new ArgumentException("The selected columns are empty or exceed the native worksheet bounds.", nameof(data));
        var projected = new DataTable(data.TableName);
        try {
            foreach (string name in names)
                projected.Columns.Add(ObjectDataHelpers.GetTableHeader(name, options),
                    options.Formatters.ContainsKey(name) ? typeof(string) : data.Columns[name]!.DataType);
            foreach (DataRow row in data.Rows)
                projected.Rows.Add(names.Select(name => ObjectDataHelpers.FormatTableValue(name,
                    row[name] is DBNull ? null : row[name], options) ?? DBNull.Value).ToArray());
            return projected;
        } catch {
            projected.Dispose();
            throw;
        }
    }
}
