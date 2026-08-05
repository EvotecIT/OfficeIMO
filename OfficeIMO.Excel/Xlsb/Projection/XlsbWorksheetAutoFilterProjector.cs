using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares the common XLSB worksheet AutoFilter subset.</summary>
    internal static class XlsbWorksheetAutoFilterProjector {
        internal static void Apply(ExcelSheet sheet, XlsbAutoFilter? source) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (source == null) return;

            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            worksheet.Append(Create(source));
        }

        internal static bool Matches(ExcelSheet sheet, XlsbAutoFilter? expected) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            AutoFilter[] actual = worksheet.Elements<AutoFilter>().ToArray();
            AutoFilter? expectedElement = expected == null ? null : Create(expected);
            return actual.Length <= 1
                && (expectedElement == null
                    ? actual.Length == 0
                    : actual.Length == 1
                        && string.Equals(actual[0].OuterXml, expectedElement.OuterXml, StringComparison.Ordinal));
        }

        private static AutoFilter Create(XlsbAutoFilter source) {
            var result = new AutoFilter { Reference = source.Range.ToA1Reference() };
            foreach (XlsbAutoFilterColumn sourceColumn in source.Columns) {
                if (sourceColumn.HasUnsupportedContent) continue;
                var column = new FilterColumn { ColumnId = sourceColumn.ColumnId };
                var filters = new Filters { Blank = sourceColumn.IncludeBlank };
                foreach (string value in sourceColumn.Values) {
                    filters.Append(new Filter { Val = value });
                }
                column.Append(filters);
                result.Append(column);
            }
            return result;
        }
    }
}
