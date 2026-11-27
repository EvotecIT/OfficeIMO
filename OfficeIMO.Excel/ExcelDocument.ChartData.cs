using System.Diagnostics;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal const string ChartDataSheetName = "OfficeIMO_ChartData";
        private const string ChartDataOwnerDefinedName = "_OfficeIMO_ChartDataOwner";
        // Created on first use rather than by field initializers, which would compile into the document's constructors.
        private object? _chartDataLockValue;
        private List<(int StartRow, int RowCount)>? _chartDataFreeRowsValue;
        private object ChartDataLock => LazyInitializer.EnsureInitialized(ref _chartDataLockValue, () => new object())!;
        private List<(int StartRow, int RowCount)> ChartDataFreeRows => LazyInitializer.EnsureInitialized(ref _chartDataFreeRowsValue)!;
        private ExcelSheet? _chartDataSheet;
        private int _chartDataNextRow;
        private const int SpreadsheetMl = OfficeXmlNamespaces.SpreadsheetML;

        // The workbook children a new definedNames element goes before, where the SDK's typed setter put it.
        private static bool FollowsDefinedNames(OfficeXmlNode child) =>
            child.NamespaceId == SpreadsheetMl && child.LocalName switch {
                "calcPr" or "oleSize" or "customWorkbookViews" or "pivotCaches" or "webPublishing" or "fileRecoveryPr"
                    or "webPublishObjects" or "extLst" => true,
                _ => false
            };

        private OfficeXmlNode WorkbookRootNode {
            get {
                OfficeXmlNode root = ExcelPartRef.Workbook(this).Root;
                return root.IsNone ? throw new InvalidOperationException("Workbook is null.") : root;
            }
        }

        // The workbook's sheet entries, in order.
        private List<OfficeXmlNode> WorkbookSheetNodes() => WorkbookRootNode.Child(SpreadsheetMl, "sheets").Children(SpreadsheetMl, "sheet").ToList();

        // The workbook-level defined names (no localSheetId) with this name.
        private static List<OfficeXmlNode> WorkbookDefinedNames(OfficeXmlNode definedNames, string name) {
            var matches = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode definedName in definedNames.Children(SpreadsheetMl, "definedName")) {
                if (definedName.Attribute("localSheetId") == null
                    && string.Equals(definedName.Attribute("name"), name, StringComparison.OrdinalIgnoreCase)) {
                    matches.Add(definedName);
                }
            }
            return matches;
        }

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
            ArgumentNullException.ThrowIfNull(sheet);
            if (rowsNeeded <= 0) rowsNeeded = 1;
            if (spacingRows < 0) spacingRows = 0;

            lock (ChartDataLock) {
                int slotRows = checked(rowsNeeded + spacingRows);
                for (int index = 0; index < ChartDataFreeRows.Count; index++) {
                    (int freeStart, int freeCount) = ChartDataFreeRows[index];
                    if (freeCount < slotRows) continue;
                    if (freeCount == slotRows) ChartDataFreeRows.RemoveAt(index);
                    else ChartDataFreeRows[index] = (freeStart + slotRows, freeCount - slotRows);
                    return freeStart;
                }
                if (_chartDataNextRow <= 0) {
                    _chartDataNextRow = CalculateInitialChartDataRow(sheet);
                }

                int startRow = _chartDataNextRow;
                long nextRow = (long)startRow + slotRows;
                if (nextRow - 1L > ExcelA1.MaxRows) {
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
            ExcelSheet dataSheet = GetSheetForLockedOperation(currentRange.SheetName);
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
                    ExcelA1.CellReference(firstRow, firstColumn) + ":" + ExcelA1.CellReference(lastRow, lastColumn),
                    ExcelClearOptions.All);
            }
            if (obsolete.Count == 1
                && obsolete[0].Equals(current)
                && !liveReferences.Any(reference => reference.Intersects(current))) {
                current.GetBounds(out int firstRow, out _, out int lastRow, out _);
                int releaseLastRow = Math.Min(ExcelA1.MaxRows, checked(lastRow + 2));
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
            List<OfficeXmlNode> sheets = WorkbookSheetNodes();
            int sheetIndex = sheets.FindIndex(sheet =>
                string.Equals(sheet.Attribute("name"), sheetName, StringComparison.OrdinalIgnoreCase));
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

        internal bool IsOwnedChartDataRangeShared(
            ExcelChartDataRange range,
            ExcelPartRef ownerPart) {
            ArgumentNullException.ThrowIfNull(range);
            if (ownerPart.IsNone) throw new ArgumentNullException(nameof(ownerPart));
            ExcelReference ownedRange = ExcelReference.Parse(range.DataRangeA1);

            IEnumerable<ExcelPartRef> chartParts = ExcelPartRef.WorksheetParts(this)
                .SelectMany(part => part.Drawings.PartsOfKind(ExcelPartKind.ExtendedChart))
                .Concat(ExcelPartRef.ChartsheetParts(this).SelectMany(part =>
                    part.Drawings.PartsOfKind(ExcelPartKind.ExtendedChart)));
            foreach (ExcelPartRef chartPart in chartParts) {
                if (chartPart == ownerPart) continue;
                foreach (string formula in EnumerateMutationFormulaTexts(chartPart.Root)) {
                    if (!ExcelChartUtils.TryParseSheetQualifiedRange(
                            formula,
                            out string sheetName,
                            out string reference)
                        || !string.Equals(sheetName, range.SheetName, StringComparison.OrdinalIgnoreCase)
                        || !ExcelReference.TryParse(reference, out ExcelReference? parsed)) {
                        continue;
                    }
                    if (parsed!.Intersects(ownedRange)) return true;
                }
            }
            return false;
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
                    ExcelA1.CellReference(r1, c1) + ":" + ExcelA1.CellReference(r2, c2)));
            }
            Add(currentFirstRow, currentFirstColumn, firstRow - 1, currentLastColumn);
            Add(lastRow + 1, currentFirstColumn, currentLastRow, currentLastColumn);
            Add(firstRow, currentFirstColumn, lastRow, firstColumn - 1);
            Add(firstRow, lastColumn + 1, lastRow, currentLastColumn);
            return ranges;
        }

        private void ReleaseChartDataRows(int startRow, int rowCount) {
            lock (ChartDataLock) {
                ChartDataFreeRows.Add((startRow, rowCount));
                ChartDataFreeRows.Sort((left, right) => left.StartRow.CompareTo(right.StartRow));
                for (int index = ChartDataFreeRows.Count - 1; index > 0; index--) {
                    (int previousStart, int previousCount) = ChartDataFreeRows[index - 1];
                    (int currentStart, int currentCount) = ChartDataFreeRows[index];
                    if (previousStart + previousCount < currentStart) continue;
                    int mergedEnd = Math.Max(previousStart + previousCount, currentStart + currentCount);
                    ChartDataFreeRows[index - 1] = (previousStart, mergedEnd - previousStart);
                    ChartDataFreeRows.RemoveAt(index);
                }
            }
        }

        private ExcelSheet GetOrCreateChartDataSheetCore() {
            if (_chartDataSheet != null) {
                return _chartDataSheet;
            }

            Stopwatch? stageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            List<OfficeXmlNode> sheets = WorkbookSheetNodes();
            ReportChartDataTiming(stageWatch, "ChartData.GetSheets");

            stageWatch?.Restart();
            string? ownedSheetName = GetOwnedChartDataSheetName();
            OfficeXmlNode existing = ownedSheetName == null
                ? OfficeXmlNode.None
                : sheets.FirstOrDefault(s => string.Equals(s.Attribute("name"), ownedSheetName, StringComparison.OrdinalIgnoreCase));
            ReportChartDataTiming(stageWatch, "ChartData.FindExistingSheet");
            if (existing.Exists) {
                stageWatch?.Restart();
                _chartDataSheet = WrapChartDataSheet(existing);
                ReportChartDataTiming(stageWatch, "ChartData.WrapExistingSheet");
                return _chartDataSheet;
            }

            stageWatch?.Restart();
            var names = new HashSet<string>(
                sheets.Select(sheet => sheet.Attribute("name") ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);
            string createdName = ChartDataSheetName;
            for (int suffix = 2; names.Contains(createdName); suffix++) {
                createdName = ChartDataSheetName + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            var created = CreateChartDataWorksheet(createdName);
            ReportChartDataTiming(stageWatch, "ChartData.CreateWorksheet");

            stageWatch?.Restart();
            using (PreserveDirectDataSetFastSaveStateDuringDirtyMarks()) {
                using (created.BeginNoLock()) {
                    created.SetHiddenWithoutSavingWorkbook(true);
                }
            }
            ReportChartDataTiming(stageWatch, "ChartData.HideWorksheet");

            OfficeXmlNode workbook = WorkbookRootNode;
            OfficeXmlNode definedNames = workbook.Child(SpreadsheetMl, "definedNames");
            if (definedNames.IsNone) {
                OfficeXmlNode follower = OfficeXmlNode.None;
                foreach (OfficeXmlNode child in workbook.Children()) {
                    if (FollowsDefinedNames(child)) {
                        follower = child;
                        break;
                    }
                }
                definedNames = workbook.InsertBefore(workbook.CreateElement(SpreadsheetMl, "definedNames"), follower);
            }
            foreach (OfficeXmlNode marker in WorkbookDefinedNames(definedNames, ChartDataOwnerDefinedName)) {
                marker.Remove();
            }
            OfficeXmlNode ownerName = definedNames.CreateElement(SpreadsheetMl, "definedName");
            ownerName.SetAttribute("name", ChartDataOwnerDefinedName);
            ownerName.SetAttribute("hidden", "1");
            ownerName.SetInnerText(ExcelChartUtils.BuildSheetQualifiedRange(created.Name, "$A$1"));
            definedNames.Append(ownerName);

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
            OfficeXmlNode ownedSheet = WorkbookSheetNodes().FirstOrDefault(sheet =>
                string.Equals(sheet.Attribute("name"), ownedSheetName, StringComparison.OrdinalIgnoreCase));
            string? state = ownedSheet.Attribute("state");
            return state == "hidden" || state == "veryHidden";
        }

        private void ValidateOwnedChartDataSheetRemoval(string sheetName) {
            if (!IsOwnedChartDataSheet(sheetName)) return;
            if (GetOwnedChartDataReferences(sheetName).Count > 0) {
                throw new InvalidOperationException(
                    $"Worksheet '{sheetName}' contains data referenced by OfficeIMO-authored charts and cannot be removed while those charts exist.");
            }
        }

        private void CompleteOwnedChartDataSheetRemoval(string sheetName) {
            string? ownedSheetName = GetOwnedChartDataSheetName();
            if (!string.Equals(ownedSheetName, sheetName, StringComparison.OrdinalIgnoreCase)) return;
            foreach (OfficeXmlNode marker in WorkbookDefinedNames(WorkbookRootNode.Child(SpreadsheetMl, "definedNames"), ChartDataOwnerDefinedName)) {
                marker.Remove();
            }
            lock (ChartDataLock) {
                _chartDataSheet = null;
                _chartDataNextRow = 0;
                ChartDataFreeRows.Clear();
            }
        }

        private string? GetOwnedChartDataSheetName() {
            OfficeXmlNode marker = WorkbookDefinedNames(WorkbookRootNode.Child(SpreadsheetMl, "definedNames"), ChartDataOwnerDefinedName)
                .FirstOrDefault(name => name.BooleanAttribute("hidden") == true);
            return ExcelChartUtils.TryParseSheetQualifiedRange(marker.IsNone ? null : marker.InnerText, out string sheetName, out string range)
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
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(used);
            if (r2 <= 1 && c2 <= 1) {
                if (!sheet.TryGetCellText(1, 1, out var text) || string.IsNullOrEmpty(text)) {
                    return 1;
                }
            }
            return checked(r2 + 3);
        }
    }
}
