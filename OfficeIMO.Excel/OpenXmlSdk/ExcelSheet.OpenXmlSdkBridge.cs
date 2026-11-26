using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Attaches the SDK's sheet element and part to a sheet the load read without the SDK.</summary>
        internal void BindToSdk(SpreadsheetDocument document, Sheet sheet, WorksheetPart part) {
            _spreadSheetDocumentValue = document;
            _sheetValue = sheet;
            _idValue = sheet.SheetId!;
            _worksheetPartValue = part;
            if (_worksheetPersistence != null) {
                // State gathered before binding belongs to the part from now on, as it would have from the start.
                WorksheetPersistenceStates.Remove(part);
                WorksheetPersistenceStates.Add(part, _worksheetPersistence);
            }
        }
    }
}
