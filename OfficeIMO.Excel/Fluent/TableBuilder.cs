using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class TableBuilder {
        private readonly ExcelSheet _sheet;
        private TableStyle _style = TableStyle.TableStyleLight9;
        private bool _hasHeader = true;
        private bool _includeAutoFilter = true;

        internal TableBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public TableBuilder Style(TableStyle style) {
            _style = style;
            return this;
        }

        public TableBuilder HasHeader(bool hasHeader = true) {
            _hasHeader = hasHeader;
            return this;
        }

        public TableBuilder WithAutoFilter(bool includeAutoFilter = true) {
            _includeAutoFilter = includeAutoFilter;
            return this;
        }

        internal void Build(string range, string name) {
            _sheet.AddTable(range, _hasHeader, name, _style, _includeAutoFilter);
        }

        public TableBuilder Add(string range, bool hasHeader = true, string name = "", TableStyle style = TableStyle.TableStyleLight9, bool includeAutoFilter = true) {
            _sheet.AddTable(range, hasHeader, name, style, includeAutoFilter);
            return this;
        }
    }
}
