using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Keep evidence attached to the cells actually written or independently qualified.
        // Structural edits move these nodes; adopting arbitrary current values would hide external edits.
        private static void RefreshDynamicSpillOwnershipAfterMutation(WorkbookPart workbook) {
            foreach (WorksheetPart part in workbook.WorksheetParts) {
                if (!DynamicSpillOwnerships.TryGetValue(part, out DynamicSpillOwnership? ownership)) continue;
                ownership.WriteIndexInitialized = false;
                var moved = new Dictionary<long, DynamicSpillCacheSnapshot>();
                foreach (DynamicSpillCacheSnapshot snapshot in ownership.Cells.Values) {
                    Cell? cell = snapshot.SourceCell;
                    if (cell?.CellReference?.Value is string reference
                        && cell.Ancestors<Worksheet>().FirstOrDefault() == part.Worksheet
                        && MatchesDynamicSpillSnapshot(cell, snapshot)
                        && A1.TryParseRange(reference, out int row, out int column, out _, out _))
                        moved[DynamicCellKey(row, column)] = snapshot;
                }
                ownership.Cells.Clear();
                foreach (var item in moved) ownership.Cells.Add(item.Key, item.Value);
                // Source XML no longer describes the structurally edited tree.
                ownership.OriginalFingerprint = null;
                ownership.OriginalScanAttempted = true;
                ownership.WrittenOwners.Clear();
            }
        }

        private static Dictionary<WorksheetPart, DynamicSpillOwnership> CaptureDynamicSpillMutationState(WorkbookPart workbook) {
            var state = new Dictionary<WorksheetPart, DynamicSpillOwnership>();
            foreach (WorksheetPart part in workbook.WorksheetParts) {
                if (!DynamicSpillOwnerships.TryGetValue(part, out DynamicSpillOwnership? ownership)) continue;
                var copy = new DynamicSpillOwnership {
                    OriginalFingerprintAttempted = ownership.OriginalFingerprintAttempted,
                    OriginalFingerprint = ownership.OriginalFingerprint,
                    OriginalScanAttempted = ownership.OriginalScanAttempted
                };
                foreach (var item in ownership.Cells)
                    copy.Cells.Add(item.Key, new DynamicSpillCacheSnapshot {
                        Type = item.Value.Type, Value = item.Value.Value, ValueMetaIndex = item.Value.ValueMetaIndex
                    });
                foreach (long key in ownership.WrittenOwners) copy.WrittenOwners.Add(key);
                state.Add(part, copy);
            }
            return state;
        }

        private static void RestoreDynamicSpillMutationState(WorkbookPart workbook,
            Dictionary<WorksheetPart, DynamicSpillOwnership> state) {
            foreach (WorksheetPart part in workbook.WorksheetParts) {
                DynamicSpillOwnerships.Remove(part);
                if (!state.TryGetValue(part, out DynamicSpillOwnership? ownership)) continue;
                foreach (Cell cell in part.Worksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>()) {
                    if (cell.CellReference?.Value is string reference
                        && A1.TryParseRange(reference, out int row, out int column, out _, out _)
                        && ownership.Cells.TryGetValue(DynamicCellKey(row, column), out DynamicSpillCacheSnapshot? snapshot)
                        && MatchesDynamicSpillSnapshot(cell, snapshot)) snapshot.SourceCell = cell;
                }
                DynamicSpillOwnerships.Add(part, ownership);
            }
        }
    }
}
