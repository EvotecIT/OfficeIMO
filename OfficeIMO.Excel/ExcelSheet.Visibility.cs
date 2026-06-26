using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets whether the worksheet is hidden or very hidden in the workbook.
        /// </summary>
        public bool Hidden => _sheet.State?.Value == SheetStateValues.Hidden || _sheet.State?.Value == SheetStateValues.VeryHidden;

        /// <summary>
        /// Gets whether the worksheet is very hidden in the workbook.
        /// </summary>
        public bool VeryHidden => _sheet.State?.Value == SheetStateValues.VeryHidden;

        /// <summary>
        /// Hides or shows the worksheet in the workbook.
        /// </summary>
        public void SetHidden(bool hidden) {
            WriteLockConditional(() => {
                SetHiddenCore(hidden);
                WorkbookRoot.Save();
            });
        }

        /// <summary>
        /// Makes the worksheet very hidden, or restores it to visible.
        /// </summary>
        /// <param name="veryHidden">Whether to set the worksheet to the very hidden workbook state.</param>
        public void SetVeryHidden(bool veryHidden) {
            WriteLockConditional(() => {
                SetVeryHiddenCore(veryHidden);
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

        private void SetVeryHiddenCore(bool veryHidden) {
            _sheet.State = veryHidden ? SheetStateValues.VeryHidden : (SheetStateValues?)null;
        }
    }
}
