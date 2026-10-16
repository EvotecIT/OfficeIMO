using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Retains a caller's worksheet handle when saving replaces the underlying package.</summary>
        internal void RebindAfterPackageReload(SpreadsheetDocument document, Sheet sheet, WorksheetPart part) {
            _spreadSheetDocument = document;
            _sheet = sheet;
            _id = sheet.SheetId!;
            _worksheetPart = part;
            _worksheetPersistence = null;
            _sheetDataCache = null;
            _rowElementCursor = null;
            _cellTextSharedStringCache = null;
            _cellValueSharedStringIndexCache = null;
            _cellValueSharedStringTable = null;
            _cellValueStyleRevision = null;
            _cellValueStyleRevisionNumber = 0;
            _cellValueDateStyleIndexes = null;
            _cellValueDurationStyleIndexes = null;
            _cellValueDefaultDateStyleIndex = null;
            _cellValueDefaultDurationStyleIndex = null;
            ResetStructuralMutationCachesLocal();
            _hasWorksheetMutations = false;
        }
    }
}
