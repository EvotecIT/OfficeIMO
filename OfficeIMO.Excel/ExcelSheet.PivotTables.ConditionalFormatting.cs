using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel worksheet.
    /// </summary>
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

        /// <summary>
        /// Adds a cell-is conditional formatting rule to a pivot table output range.
        /// </summary>
        public void AddPivotConditionalRule(string pivotTableName, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null, ExcelPivotRangeTarget target = ExcelPivotRangeTarget.DataBody) {
            AddConditionalRule(GetPivotTableRange(pivotTableName, target), @operator, formula1, formula2);
        }

        /// <summary>
        /// Adds a color scale conditional format to a pivot table output range.
        /// </summary>
        public void AddPivotConditionalColorScale(string pivotTableName, string startColor, string endColor, ExcelPivotRangeTarget target = ExcelPivotRangeTarget.DataBody) {
            AddConditionalColorScale(GetPivotTableRange(pivotTableName, target), startColor, endColor);
        }

        /// <summary>
        /// Adds a data bar conditional format to a pivot table output range.
        /// </summary>
        public void AddPivotConditionalDataBar(string pivotTableName, string color, ExcelPivotRangeTarget target = ExcelPivotRangeTarget.DataBody) {
            AddConditionalDataBar(GetPivotTableRange(pivotTableName, target), color);
        }

        /// <summary>
        /// Adds an icon set conditional format to a pivot table output range.
        /// </summary>
        public void AddPivotConditionalIconSet(string pivotTableName, IconSetValues iconSet, bool showValue, bool reverseIconOrder, double[]? percentThresholds = null, double[]? numberThresholds = null, ExcelPivotRangeTarget target = ExcelPivotRangeTarget.DataBody) {
            AddConditionalIconSet(GetPivotTableRange(pivotTableName, target), iconSet, showValue, reverseIconOrder, percentThresholds, numberThresholds);
        }

        private static string ResolvePivotDataBodyRange(string location) {
            if (!A1.TryParseRange(location, out int r1, out int c1, out int r2, out int c2)) {
                return location;
            }

            int firstRow = r1 < r2 ? r1 + 1 : r1;
            int firstColumn = c1 < c2 ? c1 + 1 : c1;
            return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(r2, c2);
        }
    }
}
