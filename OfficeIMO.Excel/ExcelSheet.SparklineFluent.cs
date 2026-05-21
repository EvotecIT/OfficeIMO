using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Starts a fluent sparkline definition from an A1 data range.
        /// </summary>
        public SparklineBuilder Sparklines(string dataRange) {
            return new SparklineBuilder(this, dataRange);
        }
    }
}
