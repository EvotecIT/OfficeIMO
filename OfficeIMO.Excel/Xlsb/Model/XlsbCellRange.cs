namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbCellRange {
        internal XlsbCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn) {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        internal int FirstRow { get; }

        internal int LastRow { get; }

        internal int FirstColumn { get; }

        internal int LastColumn { get; }

        internal string ToA1Reference() {
            string first = A1.CellReference(FirstRow, FirstColumn);
            string last = A1.CellReference(LastRow, LastColumn);
            return first == last ? first : first + ":" + last;
        }
    }
}
