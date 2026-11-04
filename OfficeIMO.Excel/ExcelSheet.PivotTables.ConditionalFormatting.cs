using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    // Represents an Excel worksheet.
    public partial class ExcelSheet {
        /// <summary>
        /// Resolves the output range for a pivot table on this worksheet.
        /// </summary>
        public string GetPivotTableRange(string pivotTableName, ExcelPivotRangeTarget target = ExcelPivotRangeTarget.WholeTable) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) {
                throw new ArgumentException("Pivot table name cannot be null or empty.", nameof(pivotTableName));
            }

            ExcelPivotTableInfo? pivot = GetPivotTables()
                .FirstOrDefault(item => string.Equals(item.Name, pivotTableName, StringComparison.OrdinalIgnoreCase));

            if (pivot == null) {
                throw new InvalidOperationException($"Pivot table '{pivotTableName}' was not found on worksheet '{Name}'.");
            }
            if (string.IsNullOrWhiteSpace(pivot.Location)) {
                throw new InvalidOperationException($"Pivot table '{pivotTableName}' does not expose an output range.");
            }

            return target == ExcelPivotRangeTarget.DataBody
                ? ResolvePivotDataBodyRange(pivot.Location!)
                : pivot.Location!;
        }

        private static string ResolvePivotDataBodyRange(string location) {
            if (!ExcelA1.TryParseRange(location, out int r1, out int c1, out int r2, out int c2)) {
                return location;
            }

            int firstRow = r1 < r2 ? r1 + 1 : r1;
            int firstColumn = c1 < c2 ? c1 + 1 : c1;
            return ExcelA1.CellReference(firstRow, firstColumn) + ":" + ExcelA1.CellReference(r2, c2);
        }
    }
}
