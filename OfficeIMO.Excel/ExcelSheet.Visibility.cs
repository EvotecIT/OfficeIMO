using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Hides or shows the worksheet in the workbook.
        /// </summary>
        public void SetHidden(bool hidden) {
            WriteLockConditional(() => {
                _sheet.State = hidden ? SheetStateValues.Hidden : (SheetStateValues?)null;
                _spreadSheetDocument.WorkbookPart?.Workbook.Save();
            });
        }
    }
}
