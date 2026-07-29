using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Starts a fluent read pipeline over this open document.
        /// </summary>
        internal ExcelFluentReadWorkbook Read(ExcelReadOptions? options = null)
            => new ExcelFluentReadWorkbook(this, options);
    }
}

