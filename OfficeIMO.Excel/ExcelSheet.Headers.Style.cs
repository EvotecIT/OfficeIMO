using System;
using System.Globalization;
using OfficeIMO.Excel.Read;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Header-based style helpers.
    /// </summary>
    public partial class ExcelSheet
    {
        

        /// <summary>
        /// Returns a builder for styling a column resolved by header with discoverable methods.
        /// </summary>
        public ColumnStyleByHeaderBuilder ColumnStyleByHeader(string header, bool includeHeader = false, ExcelReadOptions? options = null)
        {
            int colIndex = ColumnIndexByHeader(header, options);
            var a1 = GetUsedRangeA1();
            var (r1, _, r2, _) = OfficeIMO.Excel.Read.A1.ParseRange(a1);
            int startRow = includeHeader ? r1 : r1 + 1;
            return new ColumnStyleByHeaderBuilder(this, colIndex, startRow, r2);
        }
    }
}
