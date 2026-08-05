using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Validates a loaded worksheet AutoFilter edit and builds its replacement BIFF12 records.
    /// </summary>
    internal sealed class XlsbWorksheetAutoFilterPlan {
        private XlsbWorksheetAutoFilterPlan(
            bool rewrite,
            IReadOnlyList<XlsbGeneratedRecord> records) {
            Rewrite = rewrite;
            Records = records;
        }

        internal bool Rewrite { get; }

        internal IReadOnlyList<XlsbGeneratedRecord> Records { get; }

        internal static XlsbWorksheetAutoFilterPlan Create(
            ExcelSheet sheet,
            XlsbWorksheet sourceSheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));
            if (XlsbWorksheetAutoFilterProjector.Matches(sheet, sourceSheet.AutoFilter)) {
                return new XlsbWorksheetAutoFilterPlan(
                    rewrite: false,
                    Array.Empty<XlsbGeneratedRecord>());
            }

            XlsbAutoFilter? source = sourceSheet.AutoFilter;
            if (source == null) {
                throw new NotSupportedException(
                    $"Native XLSB rewriting cannot add a worksheet AutoFilter on worksheet '{sheet.Name}' yet because its workbook-level _FilterDatabase name must also be created.");
            }
            if (source.HasUnsupportedContent || source.Columns.Any(column => column.HasUnsupportedContent)) {
                throw new NotSupportedException(
                    $"Native XLSB rewriting preserves but cannot modify the worksheet AutoFilter on worksheet '{sheet.Name}' because it contains criteria outside the supported equality-list subset.");
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            AutoFilter[] filters = worksheet.Elements<AutoFilter>().ToArray();
            if (filters.Length != 1) {
                throw new NotSupportedException(
                    $"Native XLSB rewriting cannot remove the worksheet AutoFilter on worksheet '{sheet.Name}' yet because its workbook-level _FilterDatabase name must also be removed.");
            }
            AutoFilter filter = filters[0];
            if (!XlsbWorksheetAutoFilterWriter.TryGetRange(filter, out XlsbCellRange? range)
                || range == null
                || !RangesMatch(range, source.Range)) {
                throw new NotSupportedException(
                    $"Native XLSB rewriting cannot resize the worksheet AutoFilter on worksheet '{sheet.Name}' yet because its workbook-level _FilterDatabase name must also be updated.");
            }

            return new XlsbWorksheetAutoFilterPlan(
                rewrite: true,
                XlsbWorksheetAutoFilterWriter.CreateRecords(filter, sheet.Name));
        }

        private static bool RangesMatch(XlsbCellRange left, XlsbCellRange right) =>
            left.FirstRow == right.FirstRow
            && left.LastRow == right.LastRow
            && left.FirstColumn == right.FirstColumn
            && left.LastColumn == right.LastColumn;
    }
}
