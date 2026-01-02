using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string ChartDataSheetName = "OfficeIMO_ChartData";
        private readonly object _chartDataLock = new object();
        private ExcelSheet? _chartDataSheet;
        private int _chartDataNextRow;

        internal ExcelSheet GetOrCreateChartDataSheet() {
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                return GetOrCreateChartDataSheetCore();
            }

            return Locking.ExecuteWrite(EnsureLock(), GetOrCreateChartDataSheetCore);
        }

        internal int ReserveChartDataStartRow(int rowsNeeded, int spacingRows = 2) {
            if (rowsNeeded <= 0) rowsNeeded = 1;
            if (spacingRows < 0) spacingRows = 0;

            lock (_chartDataLock) {
                ExcelSheet sheet = GetOrCreateChartDataSheet();
                if (_chartDataNextRow <= 0) {
                    _chartDataNextRow = CalculateInitialChartDataRow(sheet);
                }

                int startRow = _chartDataNextRow;
                _chartDataNextRow = startRow + rowsNeeded + spacingRows;
                return startRow;
            }
        }

        private ExcelSheet GetOrCreateChartDataSheetCore() {
            if (_chartDataSheet != null) {
                return _chartDataSheet;
            }

            var sheets = _workBookPart.Workbook.Sheets?.OfType<Sheet>() ?? Enumerable.Empty<Sheet>();
            var existing = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, ChartDataSheetName, StringComparison.OrdinalIgnoreCase));
            if (existing != null) {
                _chartDataSheet = new ExcelSheet(this, _spreadSheetDocument, existing);
                return _chartDataSheet;
            }

            var created = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, ChartDataSheetName);
            using (created.BeginNoLock()) {
                created.SetHidden(true);
            }
            MarkSheetCacheDirty();
            _chartDataSheet = created;
            return created;
        }

        private static int CalculateInitialChartDataRow(ExcelSheet sheet) {
            string used = sheet.GetUsedRangeA1();
            var (r1, c1, r2, c2) = A1.ParseRange(used);
            if (r2 <= 1 && c2 <= 1) {
                if (!sheet.TryGetCellText(1, 1, out var text) || string.IsNullOrEmpty(text)) {
                    return 1;
                }
            }
            return r2 + 2;
        }
    }
}
