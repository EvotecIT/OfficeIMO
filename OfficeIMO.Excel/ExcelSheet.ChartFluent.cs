using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Starts a fluent chart definition from an A1 data range.
        /// </summary>
        public ChartBuilder Chart(string dataRangeA1) {
            return new ChartBuilder(this, dataRangeA1, isTableSource: false);
        }

        /// <summary>
        /// Starts a fluent chart definition from an existing table on this worksheet.
        /// </summary>
        public ChartBuilder ChartFromTable(string tableName) {
            return new ChartBuilder(this, tableName, isTableSource: true);
        }
    }
}
