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

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbAutoFilter? expected) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            AutoFilter[] actual = worksheet.Elements<AutoFilter>().ToArray();
            AutoFilter? expectedElement = expected == null ? null : Create(expected);
            if (actual.Length > 1
                || (expectedElement == null && actual.Length != 0)
                || (expectedElement != null
                    && (actual.Length != 1
                        || !string.Equals(actual[0].OuterXml, expectedElement.OuterXml, StringComparison.Ordinal)))) {
                throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify the worksheet AutoFilter on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
            }
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
