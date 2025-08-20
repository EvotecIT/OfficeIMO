using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ColumnBuilder {
        private readonly ExcelSheet _sheet;

        internal ColumnBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public ColumnBuilder AutoFit(bool columns = true, bool rows = false) {
            if (columns) {
                _sheet.AutoFitColumns();
            }
            if (rows) {
                _sheet.AutoFitRows();
            }
            return this;
        }
    }
}
