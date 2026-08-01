using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal const string ChartDataSheetName = "OfficeIMO_ChartData";
        private const string ChartDataOwnerDefinedName = "_OfficeIMO_ChartDataOwner";
        private readonly object _chartDataLock = new object();
        private readonly List<(int StartRow, int RowCount)> _chartDataFreeRows =
            new List<(int StartRow, int RowCount)>();
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
                int slotRows = checked(rowsNeeded + spacingRows);
                for (int index = 0; index < _chartDataFreeRows.Count; index++) {
                    (int freeStart, int freeCount) = _chartDataFreeRows[index];
                    if (freeCount < slotRows) continue;
                    if (freeCount == slotRows) _chartDataFreeRows.RemoveAt(index);
                    else _chartDataFreeRows[index] = (freeStart + slotRows, freeCount - slotRows);
                    return freeStart;
                }
                if (_chartDataNextRow <= 0) {
                    _chartDataNextRow = CalculateInitialChartDataRow(sheet);
                }

                int startRow = _chartDataNextRow;
                long nextRow = (long)startRow + slotRows;
                if (nextRow - 1L > A1.MaxRows) {
                    throw new InvalidOperationException("Chart data exceeds Excel's worksheet row limit.");
                }
                _chartDataNextRow = checked((int)nextRow);
                return startRow;
            }
        }

        internal void ReleaseOwnedChartDataRange(
            ExcelChartDataRange currentRange,
            ExcelChartDataRange? retainedRange = null) {
            if (currentRange == null || !IsOwnedChartDataSheet(currentRange.SheetName)) return;
            ExcelSheet dataSheet = this[currentRange.SheetName];
            ExcelReference current = ExcelReference.Parse(currentRange.DataRangeA1);
            ExcelReference? retained = retainedRange != null
                && string.Equals(currentRange.SheetName, retainedRange.SheetName, StringComparison.OrdinalIgnoreCase)
                    ? ExcelReference.Parse(retainedRange.DataRangeA1)
                    : null;
            IReadOnlyList<ExcelReference> liveReferences = GetOwnedChartDataReferences(currentRange.SheetName);
            IReadOnlyList<ExcelReference> obsolete = SubtractChartDataRange(current, retained);
            foreach (ExcelReference range in obsolete) {
                if (liveReferences.Any(reference => reference.Intersects(range))) continue;
                range.GetBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn);
                dataSheet.ClearRange(
                    A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn),
                    ExcelClearOptions.All);
            }
            if (obsolete.Count == 1
                && obsolete[0].Equals(current)
                && !liveReferences.Any(reference => reference.Intersects(current))) {
                current.GetBounds(out int firstRow, out _, out int lastRow, out _);
                int releaseLastRow = Math.Min(A1.MaxRows, checked(lastRow + 2));
                foreach (ExcelReference reference in liveReferences) {
                    reference.GetBounds(out int liveFirstRow, out _, out _, out _);
                    if (liveFirstRow > lastRow && liveFirstRow <= releaseLastRow) {
                        releaseLastRow = liveFirstRow - 1;
                    }
                }
                ReleaseChartDataRows(firstRow, releaseLastRow - firstRow + 1);
            }
        }

        private IReadOnlyList<ExcelReference> GetOwnedChartDataReferences(string sheetName) {
            List<Sheet> sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int sheetIndex = sheets.FindIndex(sheet =>
                string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheetIndex < 0) return Array.Empty<ExcelReference>();
            var references = new List<ExcelReference>();
            foreach (var context in EnumerateMutationFormulaContexts(
                sheets,
                sheetIndex,
                ChartDataOwnerDefinedName)) {
                foreach (ExcelFormulaReferenceSyntax node in ExcelFormulaSyntaxTree.Parse(context.Text)
                    .Nodes.OfType<ExcelFormulaReferenceSyntax>()) {
                    if (ReferenceTargetsSheet(node.Reference, sheetName, context.UnqualifiedTargetsEdited)) {
                        references.Add(node.Reference);
                    }
                }
            }
            return references;
        }

        private static IReadOnlyList<ExcelReference> SubtractChartDataRange(
            ExcelReference current,
            ExcelReference? retained) {
            current.GetBounds(out int currentFirstRow, out int currentFirstColumn, out int currentLastRow, out int currentLastColumn);
            if (retained == null || !current.Intersects(retained)) return new[] { current };
            retained.GetBounds(out int retainedFirstRow, out int retainedFirstColumn, out int retainedLastRow, out int retainedLastColumn);
            int firstRow = Math.Max(currentFirstRow, retainedFirstRow);
            int firstColumn = Math.Max(currentFirstColumn, retainedFirstColumn);
            int lastRow = Math.Min(currentLastRow, retainedLastRow);
            int lastColumn = Math.Min(currentLastColumn, retainedLastColumn);
            var ranges = new List<ExcelReference>(4);
            void Add(int r1, int c1, int r2, int c2) {
                if (r1 > r2 || c1 > c2) return;
                ranges.Add(ExcelReference.Parse(
                    A1.CellReference(r1, c1) + ":" + A1.CellReference(r2, c2)));
            }
            Add(currentFirstRow, currentFirstColumn, firstRow - 1, currentLastColumn);
            Add(lastRow + 1, currentFirstColumn, currentLastRow, currentLastColumn);
            Add(firstRow, currentFirstColumn, lastRow, firstColumn - 1);
            Add(firstRow, lastColumn + 1, lastRow, currentLastColumn);
            return ranges;
        }

        private void ReleaseChartDataRows(int startRow, int rowCount) {
            lock (_chartDataLock) {
                _chartDataFreeRows.Add((startRow, rowCount));
                _chartDataFreeRows.Sort((left, right) => left.StartRow.CompareTo(right.StartRow));
                for (int index = _chartDataFreeRows.Count - 1; index > 0; index--) {
                    (int previousStart, int previousCount) = _chartDataFreeRows[index - 1];
                    (int currentStart, int currentCount) = _chartDataFreeRows[index];
                    if (previousStart + previousCount < currentStart) continue;
                    int mergedEnd = Math.Max(previousStart + previousCount, currentStart + currentCount);
                    _chartDataFreeRows[index - 1] = (previousStart, mergedEnd - previousStart);
                    _chartDataFreeRows.RemoveAt(index);
                }
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
            return checked(r2 + 3);
        }
    }
}
