using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed partial class DirectDataSetTableModel {
            internal DataTable ToDataTable() {
                if (_sourceTable != null) {
                    return _sourceTable;
                }

                var table = new DataTable { Locale = CultureInfo.InvariantCulture };
                foreach (var column in _columns!) {
                    table.Columns.Add(column.Name, column.DataType);
                }

                table.BeginLoadData();
                try {
                    int rowCount = RowCount;
                    int columnCount = ColumnCount;
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        var values = new object?[columnCount];
                        for (int i = 0; i < values.Length; i++) {
                            values[i] = GetValue(rowIndex, i) ?? DBNull.Value;
                        }

                        table.Rows.Add(values);
                    }
                } finally {
                    table.EndLoadData();
                }

                return table;
            }
        }
    }
}
