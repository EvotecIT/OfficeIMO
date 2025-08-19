using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        public ExcelFluentWorkbook AsFluent() {
            return new ExcelFluentWorkbook(this);
        }
    }
}
