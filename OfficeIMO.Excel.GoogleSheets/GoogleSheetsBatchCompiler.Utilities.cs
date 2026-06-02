using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private static string ConvertCellToFilterText(ExcelCellSnapshot cell) {
            if (cell.Value == null) {
                return string.Empty;
            }

            return cell.Value switch {
                DateTime dateTime => dateTime.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                DateTimeOffset dateTimeOffset => dateTimeOffset.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                bool booleanValue => booleanValue ? "TRUE" : "FALSE",
                _ => Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            };
        }

        private static long GetWorksheetCellKey(int row, int column) {
            return ((long)row << 20) | (uint)column;
        }

        private static int ConvertExcelColumnWidthToPixels(double widthUnits) {
            const double mdw = 7.0;
            var pixels = Math.Truncate((256.0 * widthUnits + Math.Truncate(128.0 / mdw)) / 256.0 * mdw);
            return Math.Max(0, (int)Math.Round(pixels));
        }

        private static int ConvertPointsToPixels(double points) {
            return Math.Max(0, (int)Math.Round(points * 96.0 / 72.0));
        }
    }
}
