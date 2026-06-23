using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets the worksheet default row height in points, or null when the sheet uses the application default.
        /// </summary>
        public double? DefaultRowHeight {
            get {
                SheetFormatProperties? sheetFormat = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
                return sheetFormat?.DefaultRowHeight?.Value;
            }
        }

        /// <summary>
        /// Gets whether empty worksheet rows are hidden by default.
        /// </summary>
        public bool DefaultRowsHidden => WorksheetRoot
            .GetFirstChild<SheetFormatProperties>()
            ?.ZeroHeight
            ?.Value == true;

        /// <summary>
        /// Gets the worksheet default column width in character units, or null when the sheet uses the application default.
        /// </summary>
        public double? DefaultColumnWidth {
            get {
                SheetFormatProperties? sheetFormat = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
                return sheetFormat?.DefaultColumnWidth?.Value;
            }
        }

        /// <summary>
        /// Sets the worksheet default row height in points.
        /// </summary>
        /// <param name="height">Default row height in points. Excel supports values up to 409 points.</param>
        /// <param name="hidden">Whether empty rows should be hidden by default.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetDefaultRowHeight(double height, bool hidden = false, bool save = true) {
            if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D || height > 409D) {
                throw new ArgumentOutOfRangeException(nameof(height), "Default row height must be greater than 0 and less than or equal to 409 points.");
            }

            WriteLock(() => {
                SheetFormatProperties sheetFormat = GetOrCreateSheetFormatProperties();
                sheetFormat.DefaultRowHeight = Math.Round(height, 2);
                sheetFormat.CustomHeight = true;
                sheetFormat.ZeroHeight = hidden;
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Sets the worksheet default column width in character units.
        /// </summary>
        /// <param name="width">Default column width in character units. Excel supports values up to 255.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetDefaultColumnWidth(double width, bool save = true) {
            if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D || width > 255D) {
                throw new ArgumentOutOfRangeException(nameof(width), "Default column width must be greater than 0 and less than or equal to 255 character units.");
            }

            WriteLock(() => {
                SheetFormatProperties sheetFormat = GetOrCreateSheetFormatProperties();
                sheetFormat.DefaultColumnWidth = Math.Round(width, 2);
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        private SheetFormatProperties GetOrCreateSheetFormatProperties() {
            Worksheet worksheet = WorksheetRoot;
            SheetFormatProperties? sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
            if (sheetFormat != null) {
                return sheetFormat;
            }

            sheetFormat = new SheetFormatProperties();
            Columns? columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                worksheet.InsertBefore(sheetFormat, columns);
                return sheetFormat;
            }

            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                worksheet.InsertBefore(sheetFormat, sheetData);
                return sheetFormat;
            }

            worksheet.Append(sheetFormat);
            return sheetFormat;
        }
    }
}
