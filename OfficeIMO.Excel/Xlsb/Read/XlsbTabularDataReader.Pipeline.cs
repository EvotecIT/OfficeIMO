using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        /// <summary>Proves the same full worksheet while the package reader publishes initialized prefixes.</summary>
        internal static XlsbWorksheetDiscovery DiscoverWorksheet(XlsbPooledPartStream part,
            XlsbImportOptions limits, XlsbRecordReadBudget records, XlsbCellReadBudget cells,
            ExcelSerialDateStyle[] styles, int sharedStringCount, ExcelReaderOptions options,
            CancellationToken cancellationToken, WorksheetInflationProgress progress) {
            DiscoverDataColumns(part, limits, records, cells, styles, sharedStringCount,
                options.TreatDatesUsingNumberFormat, options.UseCachedFormulaResult, cancellationToken,
                out int firstColumn, out int lastColumn, out int firstRow, out int lastRow,
                out XlsbValidatedRowPlan? rowPlan, progress);
            return new XlsbWorksheetDiscovery(firstColumn, lastColumn, firstRow, lastRow, rowPlan);
        }
    }

    /// <summary>Retains the validation result without keeping a second worksheet representation.</summary>
    internal readonly struct XlsbWorksheetDiscovery {
        internal XlsbWorksheetDiscovery(int firstColumn, int lastColumn, int firstRow,
            int lastRow, XlsbValidatedRowPlan? rowPlan) {
            FirstColumn = firstColumn; LastColumn = lastColumn;
            FirstDataRow = firstRow; LastDataRow = lastRow; RowPlan = rowPlan;
        }
        internal int FirstColumn { get; }
        internal int LastColumn { get; }
        internal int FirstDataRow { get; }
        internal int LastDataRow { get; }
        internal XlsbValidatedRowPlan? RowPlan { get; }
    }
}
