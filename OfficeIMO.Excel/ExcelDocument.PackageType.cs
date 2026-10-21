using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Preserves unsaved workbook metadata and live worksheet identities when the SDK replaces its main part.</summary>
        private void ChangeSpreadsheetDocumentType(SpreadsheetDocumentType target) {
            if (_spreadSheetDocument.DocumentType == target) return;

            var previousParts = CaptureActiveSheetPartsForReload();
            var drawingHandles = CaptureDrawingHandlesForReload(previousParts);
            // The SDK copies persisted main-part bytes, so save its DOM before replacing it.
            WorkbookPartRoot.Workbook?.Save();
            _spreadSheetDocument.ChangeDocumentType(target);
            _workBookPart = _spreadSheetDocument.WorkbookPart
                ?? throw new InvalidOperationException("WorkbookPart is missing after the document type change.");
            _cachedSheets = null;
            _sheetCacheDirty = true;
            RebindSheetWrappersAfterPackageReload(previousParts, packageWasSaved: false);
            RebindDrawingHandlesAfterPackageReload(drawingHandles);
        }
    }
}
