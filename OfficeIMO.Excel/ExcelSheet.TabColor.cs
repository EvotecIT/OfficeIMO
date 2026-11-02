using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

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

            SetTabColor(OfficeColor.Parse(color));
        }

        /// <summary>Sets a literal or theme-based worksheet tab color.</summary>
        /// <param name="color">A literal RGBA color, native theme role or status color.</param>
        /// <remarks>Native theme roles follow later workbook theme edits; status roles are written as literals.</remarks>
        public void SetTabColor(OfficeColor color) {
            uint? themeIndex = OfficeOpenXmlThemeColorResolver.GetSpreadsheetThemeColorIndex(color.ThemeSlot);
            TabColor native = themeIndex.HasValue
                ? new TabColor { Theme = themeIndex.Value }
                : new TabColor { Rgb = color.Resolve(_excelDocument.EffectiveCreationTheme).ToArgbHex() };
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
