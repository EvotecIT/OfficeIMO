using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly object _sheetWrapperCacheLock = new object();
        private readonly List<WeakReference<ExcelSheet>> _sheetWrappers = new();
        private Dictionary<uint, WorksheetPart>? _sheetPartsBeforePackageClose;

        internal void RegisterSheetWrapper(ExcelSheet sheet) {
            lock (_sheetWrapperCacheLock) {
                _sheetWrappers.Add(new WeakReference<ExcelSheet>(sheet));
            }
        }

        /// <summary>Captures package-part identity before a save replaces the package.</summary>
        private Dictionary<uint, WorksheetPart> CaptureActiveSheetPartsForReload() {
            // Changing XLSX/template/macro package type can replace the WorkbookPart before closing.
            WorkbookPart workbookPart = WorkbookPartRoot;
            return (workbookPart.Workbook?.Sheets?.OfType<Sheet>() ?? Enumerable.Empty<Sheet>())
                .Where(sheet => sheet.SheetId != null && sheet.Id?.Value != null
                    && workbookPart.GetPartById(sheet.Id.Value) is WorksheetPart)
                .ToDictionary(sheet => sheet.SheetId!.Value, sheet => (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!));
        }

        /// <summary>Captures active worksheet identities while package relationships can still be read.</summary>
        private void ClosePackageForReload() {
            _sheetPartsBeforePackageClose = CaptureActiveSheetPartsForReload();
            _drawingHandlesBeforePackageClose = CaptureDrawingHandlesForReload(_sheetPartsBeforePackageClose);
            try { _spreadSheetDocument.Dispose(); } catch { }
        }

        /// <summary>Reconnects all live worksheet handles, including uncached aliases, after a save reloads the package.</summary>
        private void RebindSheetWrappersAfterPackageReload(IReadOnlyDictionary<uint, WorksheetPart> previousParts, bool packageWasSaved = true) {
            var elements = ReadSheetElements().Where(sheet => sheet.SheetId != null)
                .ToDictionary(sheet => sheet.SheetId!.Value);
            lock (_sheetWrapperCacheLock) {
                for (int index = _sheetWrappers.Count - 1; index >= 0; index--) {
                    if (!_sheetWrappers[index].TryGetTarget(out ExcelSheet? wrapper)) {
                        _sheetWrappers.RemoveAt(index);
                        continue;
                    }

                    Sheet previous = wrapper.SheetElement;
                    // An attached old Sheet element alone does not prove active identity after rollback.
                    // A removed handle must never attach to a replacement reusing its ID.
                    if (previous.SheetId == null
                        || !previousParts.TryGetValue(previous.SheetId.Value, out WorksheetPart? previousPart)
                        || !ReferenceEquals(wrapper.DeferredMetadataWorksheetPart, previousPart)
                        || !elements.TryGetValue(previous.SheetId.Value, out Sheet? sheet)
                        || sheet.Id?.Value == null
                        || _workBookPart.GetPartById(sheet.Id.Value) is not WorksheetPart part) continue;

                    wrapper.RebindAfterPackageReload(_spreadSheetDocument!, sheet, part, packageWasSaved);
                }
            }
        }

        internal int RegisteredSheetWrapperCountForTests {
            get {
                lock (_sheetWrapperCacheLock) {
                    return _sheetWrappers.Count;
                }
            }
        }

        internal void ResetStructuralMutationCaches(WorksheetPart worksheetPart) {
            var wrappers = new List<ExcelSheet>();
            lock (_sheetWrapperCacheLock) {
                for (int index = _sheetWrappers.Count - 1; index >= 0; index--) {
                    if (!_sheetWrappers[index].TryGetTarget(out ExcelSheet? sheet)) {
                        _sheetWrappers.RemoveAt(index);
                        continue;
                    }

                    if (ReferenceEquals(sheet.DeferredMetadataWorksheetPart, worksheetPart)) {
                        wrappers.Add(sheet);
                    }
                }
            }

            foreach (ExcelSheet sheet in wrappers) {
                sheet.ResetStructuralMutationCachesLocal();
            }
        }
    }
}
