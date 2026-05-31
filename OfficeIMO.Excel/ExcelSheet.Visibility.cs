using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets whether the worksheet is hidden or very hidden in the workbook.
        /// </summary>
        public bool Hidden => _sheet.State?.Value == SheetStateValues.Hidden || _sheet.State?.Value == SheetStateValues.VeryHidden;

        /// <summary>
        /// Hides or shows the worksheet in the workbook.
        /// </summary>
        public void SetHidden(bool hidden) {
            WriteLockConditional(() => {
                SetHiddenCore(hidden);
                WorkbookRoot.Save();
            });
        }

        internal void SetHiddenWithoutSavingWorkbook(bool hidden) {
            if (Locking.IsNoLock) {
                SetHiddenCore(hidden);
                MarkRequiresSavePreparation();
                return;
            }

            WriteLockConditional(() => SetHiddenCore(hidden));
        }

        private void SetHiddenCore(bool hidden) {
            _sheet.State = hidden ? SheetStateValues.Hidden : (SheetStateValues?)null;
        }
    }
}
