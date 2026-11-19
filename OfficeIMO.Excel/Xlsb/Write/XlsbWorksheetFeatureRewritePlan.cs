using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Compares a loaded worksheet's feature elements with the ones decoded from its source records. Unchanged collections
    /// keep their native records; changed ones are re-encoded with the same writers new workbooks use.
    /// </summary>
    internal static class XlsbWorksheetFeatureRewritePlan {
        internal static IReadOnlyList<XlsbRecordBlockEdit> Create(XlsbWorksheetFeatureContext context, XlsbSheet source) {
            ArgumentNullException.ThrowIfNull(context);
            ArgumentNullException.ThrowIfNull(source);
            var edits = new List<XlsbRecordBlockEdit>();
            if (!XlsbWorksheetFeatureProjector.Matches<RowBreaks>(context.Sheet, source)
                || !XlsbWorksheetFeatureProjector.Matches<ColumnBreaks>(context.Sheet, source)) {
                edits.Add(new XlsbRecordBlockEdit(XlsbWorksheetPageBreakWriter.BrtBeginRwBrk, XlsbWorksheetPageBreakWriter.BrtEndRwBrk,
                    XlsbWorksheetPageBreakWriter.CreateRowBreaks(context)));
                edits.Add(new XlsbRecordBlockEdit(XlsbWorksheetPageBreakWriter.BrtBeginColBrk, XlsbWorksheetPageBreakWriter.BrtEndColBrk,
                    XlsbWorksheetPageBreakWriter.CreateColumnBreaks(context)));
            }
            if (!XlsbWorksheetFeatureProjector.Matches<IgnoredErrors>(context.Sheet, source)) {
                edits.Add(new XlsbRecordBlockEdit(XlsbWorksheetIgnoredErrorWriter.BrtBeginCellIgnoreECs,
                    XlsbWorksheetIgnoredErrorWriter.BrtEndCellIgnoreECs, XlsbWorksheetIgnoredErrorWriter.Create(context)));
            }
            if (!XlsbWorksheetFeatureProjector.Matches<DataValidations>(context.Sheet, source)) {
                if (source.HasUndecodedDataValidations) {
                    throw new NotSupportedException($"Native XLSB rewrite cannot change the data validations of worksheet '{context.Sheet.Name}' because some of its existing validations could not be decoded.");
                }
                edits.Add(new XlsbRecordBlockEdit(XlsbWorksheetDataValidationWriter.BrtBeginDVals,
                    XlsbWorksheetDataValidationWriter.BrtEndDVals, XlsbWorksheetDataValidationWriter.Create(context)));
            }
            return edits;
        }
    }
}
