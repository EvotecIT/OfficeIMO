namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an SBaseRef PivotTable-view range referenced by a chart.
    /// </summary>
    public sealed class LegacyXlsChartPivotViewReference {
        internal LegacyXlsChartPivotViewReference(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
        }

        /// <summary>Gets the one-based first row of the referenced PivotTable view.</summary>
        public int FirstRow { get; }

        /// <summary>Gets the one-based first column of the referenced PivotTable view.</summary>
        public int FirstColumn { get; }

        /// <summary>Gets the one-based last row of the referenced PivotTable view.</summary>
        public int LastRow { get; }

        /// <summary>Gets the one-based last column of the referenced PivotTable view.</summary>
        public int LastColumn { get; }

        /// <summary>Gets the referenced PivotTable-view range in A1 notation.</summary>
        public string Reference {
            get {
                string start = A1.CellReference(FirstRow, FirstColumn);
                string end = A1.CellReference(LastRow, LastColumn);
                return start == end ? start : start + ":" + end;
            }
        }
    }
}
