using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ColumnBuilder {
        private readonly ExcelSheet _sheet;

        internal ColumnBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public ColumnBuilder AutoFit() {
            _sheet.AutoFitColumns();
            return this;
        }
    }
}
