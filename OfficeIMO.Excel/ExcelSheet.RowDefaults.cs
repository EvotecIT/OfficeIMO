using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets the worksheet default row height as a physical dimension, or null when the sheet uses the application default.
        /// </summary>
        public ExcelColumnWidth? DefaultRowHeight {
            get {
                SheetFormatProperties? sheetFormat = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
                return sheetFormat?.DefaultRowHeight?.Value is double height ? ExcelColumnWidth.Length(OfficeLength.Points(height)) : (ExcelColumnWidth?)null;
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
        public ExcelColumnWidth? DefaultColumnWidth {
            get {
                SheetFormatProperties? sheetFormat = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
                return sheetFormat?.DefaultColumnWidth?.Value is double width ? ExcelColumnWidth.Characters(width) : (ExcelColumnWidth?)null;
            }
        }

        /// <summary>Gets the unpadded base column width used when no default width is authored.</summary>
        internal double BaseColumnWidth => WorksheetRoot.GetFirstChild<SheetFormatProperties>()
            ?.BaseColumnWidth?.Value ?? 8U;

        /// <summary>
        /// Sets the worksheet default row height from characters or a physical length.
        /// </summary>
        /// <param name="height">Default row height. Excel supports physical heights up to 409 points.</param>
        /// <param name="hidden">Whether empty rows should be hidden by default.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetDefaultRowHeight(ExcelColumnWidth height, bool hidden = false, bool save = true) =>
            SetDefaultRowHeightCore(height.ToRowPoints(this), hidden, save, roundToHundredths: true);

        internal void SetDefaultRowHeightExact(double height, bool hidden = false, bool save = true) =>
            SetDefaultRowHeightCore(height, hidden, save, roundToHundredths: false);

        private void SetDefaultRowHeightCore(double height, bool hidden, bool save,
            bool roundToHundredths) {
            if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D || height > 409D) {
                throw new ArgumentOutOfRangeException(nameof(height), "Default row height must be greater than 0 and less than or equal to 409 points.");
            }

            WriteLock(() => {
                SheetFormatProperties sheetFormat = GetOrCreateSheetFormatProperties();
                sheetFormat.DefaultRowHeight = roundToHundredths ? Math.Round(height, 2) : height;
                sheetFormat.CustomHeight = true;
                sheetFormat.ZeroHeight = hidden;
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Sets the worksheet default column width from characters or a physical length.
        /// </summary>
        /// <param name="width">Default column width. Excel supports resolved widths up to 255 character units.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetDefaultColumnWidth(ExcelColumnWidth width, bool save = true) =>
            SetDefaultColumnWidthCore(width.ToCharacters(this), save, roundToHundredths: true);

        internal void SetDefaultColumnWidthExact(double width, bool save = true) =>
            SetDefaultColumnWidthCore(width, save, roundToHundredths: false);

        private void SetDefaultColumnWidthCore(double width, bool save,
            bool roundToHundredths) {
            if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D || width > 255D) {
                throw new ArgumentOutOfRangeException(nameof(width), "Default column width must be greater than 0 and less than or equal to 255 character units.");
            }

            WriteLock(() => {
                SheetFormatProperties sheetFormat = GetOrCreateSheetFormatProperties();
                sheetFormat.DefaultColumnWidth = roundToHundredths ? Math.Round(width, 2) : width;
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
