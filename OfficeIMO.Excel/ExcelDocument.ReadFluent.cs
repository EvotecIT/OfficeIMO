using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel
{
    public partial class ExcelDocument
    {
        /// <summary>
        /// Starts a fluent read pipeline over this open document.
        /// </summary>
        public ExcelFluentReadWorkbook Read(ExcelReadOptions? options = null)
            => new ExcelFluentReadWorkbook(this, options);
    }
}

