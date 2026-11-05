using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Sets a literal or theme-based worksheet tab color.</summary>
        /// <param name="color">A literal RGBA color, native theme role or status color.</param>
        /// <remarks>Native theme roles follow later workbook theme edits; status roles are written as literals.</remarks>
        public void SetTabColor(OfficeColor color) {
            Color resolved = ExcelColor.ToSpreadsheetColor(color, _excelDocument.ColorTheme);
            TabColor native = new TabColor { Theme = resolved.Theme, Rgb = resolved.Rgb };
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetProperties sheetProperties = WorksheetRoot.GetFirstChild<SheetProperties>() ?? new SheetProperties();
                if (sheetProperties.Parent == null) {
                    WorksheetRoot.InsertAt(sheetProperties, 0);
                }

                sheetProperties.TabColor = native;
            });
        }

        /// <summary>
        /// Clears the worksheet tab color.
        /// </summary>
        public void ClearTabColor() {
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetProperties? sheetProperties = WorksheetRoot.GetFirstChild<SheetProperties>();
                if (sheetProperties == null) {
                    return;
                }

                sheetProperties.TabColor = null;
                if (!sheetProperties.HasChildren && !sheetProperties.HasAttributes) {
                    sheetProperties.Remove();
                }

            });
        }
    }
}
