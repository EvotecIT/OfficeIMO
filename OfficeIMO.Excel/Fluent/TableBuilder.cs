using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class TableBuilder {
        private readonly ExcelSheet _sheet;

        internal TableBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public TableBuilder Add(string range, bool hasHeader = true, string name = "", TableStyle style = TableStyle.TableStyleLight9) {
            _sheet.AddTable(range, hasHeader, name, style);
            return this;
        }
    }
}
