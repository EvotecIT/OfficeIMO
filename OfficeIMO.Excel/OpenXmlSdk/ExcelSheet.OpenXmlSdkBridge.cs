using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Attaches the SDK's sheet element and part to a sheet the load read without the SDK.</summary>
        /// <summary>A sheet a new workbook added before its SDK opened is a new worksheet to the model, as one it adds is.</summary>
        internal void MarkAsNewWorksheet() {
            _hasWorksheetMutations = true;
            _requiresSavePreparation = true;
        }

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
