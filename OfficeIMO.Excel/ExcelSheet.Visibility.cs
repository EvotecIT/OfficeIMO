using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
