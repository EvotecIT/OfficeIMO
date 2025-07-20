using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single cell within a worksheet.
    /// </summary>
    public class ExcelCell {
        internal readonly Cell _cell;
        private readonly ExcelSheet _sheet;

        internal ExcelCell(ExcelSheet sheet, Cell cell) {
            _sheet = sheet;
            _cell = cell;
        }

        /// <summary>
        /// Gets or sets the text value of the cell.
        /// </summary>
        public string? Text {
            get => _cell.CellValue?.InnerText;
            set {
                if (_cell.DataType == null) {
                    _cell.DataType = CellValues.InlineString;
                }
                _cell.CellValue = new CellValue(value);
            }
        }

        /// <summary>
        /// Gets or sets the font applied to the cell.
        /// </summary>
        public Font? Font {
            get => _sheet.GetFont(this);
            set => _sheet.ApplyStyle(this, value, Fill, Border, NumberFormat);
        }

        /// <summary>
        /// Gets or sets the fill applied to the cell.
        /// </summary>
        public Fill? Fill {
            get => _sheet.GetFill(this);
            set => _sheet.ApplyStyle(this, Font, value, Border, NumberFormat);
        }

        /// <summary>
        /// Gets or sets the border applied to the cell.
        /// </summary>
        public Border? Border {
            get => _sheet.GetBorder(this);
            set => _sheet.ApplyStyle(this, Font, Fill, value, NumberFormat);
        }

        /// <summary>
        /// Gets or sets the number format applied to the cell.
        /// </summary>
        public string? NumberFormat {
            get => _sheet.GetNumberFormat(this);
            set => _sheet.ApplyStyle(this, Font, Fill, Border, value);
        }
    }
}
