using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;

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
                out XlsbValidatedRowPlan? rowPlan, out bool inlineTextFreeOfSurrogates, progress);
            return new XlsbWorksheetDiscovery(firstColumn, lastColumn, firstRow, lastRow, rowPlan,
                inlineTextFreeOfSurrogates);
        }

        /// <summary>Runs the same full validation for a single-pass reader, publishing each completed row.</summary>
        internal static void DiscoverWorksheetForSinglePass(XlsbPooledPartStream part,
            XlsbImportOptions limits, XlsbRecordReadBudget records, XlsbCellReadBudget cells,
            ExcelSerialDateStyle[] styles, int sharedStringCount, ExcelReaderOptions options,
            WorksheetInflationProgress progress, XlsbSinglePassGate gate) {
            try {
                DiscoverDataColumns(part, limits, records, cells, styles, sharedStringCount,
                    options.TreatDatesUsingNumberFormat, options.UseCachedFormulaResult, gate.Token,
                    out int firstColumn, out int lastColumn, out int firstRow, out int lastRow,
                    out XlsbValidatedRowPlan? rowPlan, out bool inlineTextFreeOfSurrogates, progress, gate);
                gate.Complete(new XlsbWorksheetDiscovery(firstColumn, lastColumn, firstRow, lastRow, rowPlan,
                    inlineTextFreeOfSurrogates), part.DataLength);
            } catch (Exception exception) {
                gate.Fail(exception);
            }
        }
    }

    /// <summary>Retains the validation result without keeping a second worksheet representation.</summary>
    internal readonly struct XlsbWorksheetDiscovery {
        internal XlsbWorksheetDiscovery(int firstColumn, int lastColumn, int firstRow,
            int lastRow, XlsbValidatedRowPlan? rowPlan, bool inlineTextFreeOfSurrogates) {
            FirstColumn = firstColumn; LastColumn = lastColumn;
            FirstDataRow = firstRow; LastDataRow = lastRow; RowPlan = rowPlan;
            InlineTextFreeOfSurrogates = inlineTextFreeOfSurrogates;
        }
        internal int FirstColumn { get; }
        internal int LastColumn { get; }
        internal int FirstDataRow { get; }
        internal int LastDataRow { get; }
        internal XlsbValidatedRowPlan? RowPlan { get; }
        internal bool InlineTextFreeOfSurrogates { get; }
    }
}
