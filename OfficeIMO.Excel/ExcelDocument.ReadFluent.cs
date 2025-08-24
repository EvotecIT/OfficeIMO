using OfficeIMO.Excel.Read;
using OfficeIMO.Excel.Read.Fluent;

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

