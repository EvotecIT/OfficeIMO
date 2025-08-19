using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class StyleBuilder {
        private readonly ExcelSheet _sheet;

        internal StyleBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public StyleBuilder SetCellFormat(int row, int column, string numberFormat) {
            _sheet.SetCellFormat(row, column, numberFormat);
            return this;
        }
    }
}
