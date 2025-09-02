using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Returns a fluent API wrapper for this document.
        /// </summary>
        public ExcelFluentWorkbook AsFluent() {
            return new ExcelFluentWorkbook(this);
        }
    }
}
