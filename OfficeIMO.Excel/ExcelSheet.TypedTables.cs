using System.Data;
using System.Diagnostics.CodeAnalysis;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an editable native table from typed properties, fields, or string-keyed dictionary columns.</summary>
    public ExcelTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null, int row = 1, int column = 1) {
        ValidateTypedTableOrigin(row, column);
        bool headers = options?.IncludeHeaders ?? true;
        var data = ObjectDataHelpers.ProjectTypedTable(rows, headers,
            ExcelA1.MaxRows - row + 1 - (headers ? 1 : 0), ExcelA1.MaxColumns - column + 1);
        int count = data.Rows.Count + (headers ? 1 : 0);
        if (count == 0) throw new ArgumentException("A headerless table requires at least one row.", nameof(rows));
        if (headers)
            for (int index = 0; index < data.Headers.Length; index++) CellValue(row, column + index, data.Headers[index]);
        for (int index = 0; index < data.Rows.Count; index++)
            for (int field = 0; field < data.Headers.Length; field++) CellValue(row + index + (headers ? 1 : 0), column + field, data.Rows[index][field]);
        string range = ExcelA1.CellReference(row, column) + ":" + ExcelA1.CellReference(row + count - 1, column + data.Headers.Length - 1);
        return AddTable(range, headers, options?.Name ?? "Table", ExcelTableStyle.TableStyleMedium2);
    }

    /// <summary>Adds an editable native table from a data table, preserving its typed cell values.</summary>
    public ExcelTable AddTable(DataTable data, OfficeTableOptions? options = null, int row = 1, int column = 1) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        ValidateTypedTableOrigin(row, column);
        bool headers = options?.IncludeHeaders ?? true;
        int count = checked(data.Rows.Count + (headers ? 1 : 0));
        if (count == 0 || data.Columns.Count == 0 || count > ExcelA1.MaxRows - row + 1 || data.Columns.Count > ExcelA1.MaxColumns - column + 1)
            throw new ArgumentException("The data table is empty or exceeds the native worksheet bounds.", nameof(data));
        InsertDataTable(data, row, column, headers);
        string range = ExcelA1.CellReference(row, column) + ":" + ExcelA1.CellReference(row + count - 1, column + data.Columns.Count - 1);
        return AddTable(range, headers, options?.Name ?? data.TableName, ExcelTableStyle.TableStyleMedium2);
    }

    private static void ValidateTypedTableOrigin(int row, int column) {
        if (row < 1 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
    }
}
