using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Sets the worksheet tab color shown by spreadsheet applications.
        /// </summary>
        /// <param name="color">Named color, RGB hex, or RGBA hex value.</param>
        public void SetTabColor(string color) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Tab color is required.", nameof(color));
            }

            OfficeColor parsed = OfficeColor.Parse(color);
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetProperties sheetProperties = WorksheetRoot.GetFirstChild<SheetProperties>() ?? new SheetProperties();
                if (sheetProperties.Parent == null) {
                    WorksheetRoot.InsertAt(sheetProperties, 0);
                }

                sheetProperties.TabColor = new TabColor {
                    Rgb = parsed.ToArgbHex()
                };
                WorksheetRoot.Save();
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

                WorksheetRoot.Save();
            });
        }
    }
}
