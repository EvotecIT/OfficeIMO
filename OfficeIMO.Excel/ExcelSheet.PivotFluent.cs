using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Starts a fluent pivot table definition from an A1 source range.
        /// </summary>
        /// <param name="sourceRange">Source data range including headers, for example A1:D100.</param>
        public PivotTableBuilder Pivot(string sourceRange) {
            return new PivotTableBuilder(this, sourceRange);
        }
    }
}
