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
        /// <remarks>
        /// Hiding the active worksheet activates the nearest visible sheet, as Excel does, when another sheet is visible. Excel shows a
        /// workbook whose active tab is hidden by unhiding that tab.
        /// </remarks>
        public void SetHidden(bool hidden) {
            WriteLockConditional(() => {
                SetHiddenCore(hidden);
                if (hidden) _excelDocument.MoveActiveTabsOffHiddenSheetsCore();
                WorkbookRoot.Save();
            });
        }

        /// <summary>
        /// Makes the worksheet very hidden, or restores it to visible.
        /// </summary>
        /// <param name="veryHidden">Whether to set the worksheet to the very hidden workbook state.</param>
        /// <remarks>Making the active worksheet very hidden activates the nearest visible sheet, as <see cref="SetHidden"/> does.</remarks>
        public void SetVeryHidden(bool veryHidden) {
            WriteLockConditional(() => {
                SetVeryHiddenCore(veryHidden);
                if (veryHidden) _excelDocument.MoveActiveTabsOffHiddenSheetsCore();
                WorkbookRoot.Save();
            });
        }

        /// <summary>
        /// Applies a sheet state read from a source workbook. The source's own workbook view chooses the active
        /// sheet, so this does not move the activation.
        /// </summary>
        internal void ApplyImportedVisibility(bool hidden, bool veryHidden) {
            WriteLockConditional(() => {
                _sheet.State = veryHidden ? SheetStateValues.VeryHidden : hidden ? SheetStateValues.Hidden : (SheetStateValues?)null;
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
