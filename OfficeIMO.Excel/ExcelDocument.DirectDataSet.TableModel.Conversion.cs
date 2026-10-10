using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed partial class DirectDataSetTableModel {
#if NET5_0_OR_GREATER
            [UnconditionalSuppressMessage(
                "Trimming",
                "IL2072",
                Justification = "OfficeIMO preserves the DataTable column type supplied by the caller. Framework scalar types used by the AOT-safe tabular APIs are statically rooted; applications that use custom DataColumn types must preserve those model members at the application boundary.")]
#endif
            internal DataTable ToDataTable(bool preserveMissingValues = true) {
                if (_sourceTable != null && preserveMissingValues) {
                    return _sourceTable;
                }

                var table = new DataTable { Locale = CultureInfo.InvariantCulture };
                for (int i = 0; i < ColumnCount; i++) {
                    table.Columns.Add(GetColumnName(i), preserveMissingValues ? GetColumnType(i) : typeof(object));
                }

                table.BeginLoadData();
                try {
                    int rowCount = RowCount;
                    int columnCount = ColumnCount;
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        var values = new object?[columnCount];
                        for (int i = 0; i < values.Length; i++) {
                            values[i] = GetValue(rowIndex, i) ?? (preserveMissingValues ? DBNull.Value : string.Empty);
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
