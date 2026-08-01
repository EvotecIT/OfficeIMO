using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal const string ChartDataSheetName = "OfficeIMO_ChartData";
        private const string ChartDataOwnerDefinedName = "_OfficeIMO_ChartDataOwner";
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
            ExcelSheet sheet = GetOrCreateChartDataSheet();
            return ReserveChartDataStartRow(sheet, rowsNeeded, spacingRows);
        }

        internal int ReserveChartDataStartRow(ExcelSheet sheet, int rowsNeeded, int spacingRows = 2) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (rowsNeeded <= 0) rowsNeeded = 1;
            if (spacingRows < 0) spacingRows = 0;

            lock (_chartDataLock) {
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

            Stopwatch? stageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            var sheets = WorkbookRoot.Sheets?.OfType<Sheet>() ?? Enumerable.Empty<Sheet>();
            ReportChartDataTiming(stageWatch, "ChartData.GetSheets");

            stageWatch?.Restart();
            string? ownedSheetName = GetOwnedChartDataSheetName();
            var existing = ownedSheetName == null
                ? null
                : sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, ownedSheetName, StringComparison.OrdinalIgnoreCase));
            ReportChartDataTiming(stageWatch, "ChartData.FindExistingSheet");
            if (existing != null) {
                stageWatch?.Restart();
                _chartDataSheet = new ExcelSheet(this, _spreadSheetDocument, existing);
                ReportChartDataTiming(stageWatch, "ChartData.WrapExistingSheet");
                return _chartDataSheet;
            }

            stageWatch?.Restart();
            var names = new HashSet<string>(
                sheets.Select(sheet => sheet.Name?.Value ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);
            string createdName = ChartDataSheetName;
            for (int suffix = 2; names.Contains(createdName); suffix++) {
                createdName = ChartDataSheetName + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            var created = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, createdName);
            ReportChartDataTiming(stageWatch, "ChartData.CreateWorksheet");

            stageWatch?.Restart();
            using (PreserveDirectDataSetFastSaveStateDuringDirtyMarks()) {
                using (created.BeginNoLock()) {
                    created.SetHiddenWithoutSavingWorkbook(true);
                }
            }
            ReportChartDataTiming(stageWatch, "ChartData.HideWorksheet");

            DefinedNames definedNames = WorkbookRoot.DefinedNames ??= new DefinedNames();
            foreach (DefinedName marker in definedNames.Elements<DefinedName>().Where(name =>
                name.LocalSheetId == null
                && string.Equals(name.Name?.Value, ChartDataOwnerDefinedName, StringComparison.OrdinalIgnoreCase)).ToList()) {
                marker.Remove();
            }
            definedNames.Append(new DefinedName {
                Name = ChartDataOwnerDefinedName,
                Hidden = true,
                Text = ExcelChartUtils.BuildSheetQualifiedRange(created.Name, "$A$1")
            });

            stageWatch?.Restart();
            using (PreserveDirectDataSetFastSaveStateDuringDirtyMarks()) {
                MarkSheetCacheDirty();
            }
            ReportChartDataTiming(stageWatch, "ChartData.MarkSheetCacheDirty");
            _chartDataSheet = created;
            _chartDataNextRow = 1;
            return created;
        }

        internal bool IsOwnedChartDataSheet(string sheetName) {
            string? ownedSheetName = GetOwnedChartDataSheetName();
            if (!string.Equals(ownedSheetName, sheetName, StringComparison.OrdinalIgnoreCase)) return false;
            return Sheets.Any(sheet => string.Equals(sheet.Name, ownedSheetName, StringComparison.OrdinalIgnoreCase) && sheet.Hidden);
        }

        private string? GetOwnedChartDataSheetName() {
            DefinedName? marker = WorkbookRoot.DefinedNames?.Elements<DefinedName>().FirstOrDefault(name =>
                name.LocalSheetId == null
                && name.Hidden?.Value == true
                && string.Equals(name.Name?.Value, ChartDataOwnerDefinedName, StringComparison.OrdinalIgnoreCase));
            return ExcelChartUtils.TryParseSheetQualifiedRange(marker?.Text, out string sheetName, out string range)
                && string.Equals(range, "A1", StringComparison.OrdinalIgnoreCase)
                ? sheetName
                : null;
        }

        private void ReportChartDataTiming(Stopwatch? stopwatch, string operation) {
            if (stopwatch != null) {
                Execution.ReportTiming(operation, stopwatch.Elapsed);
            }
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
