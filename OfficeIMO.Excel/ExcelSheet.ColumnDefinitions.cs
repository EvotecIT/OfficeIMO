using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets explicit worksheet column definitions such as custom widths and hidden ranges.
        /// </summary>
        public IReadOnlyList<ExcelColumnSnapshot> GetColumnDefinitions() {
            Columns? columns = WorksheetRoot.GetFirstChild<Columns>();
            if (columns == null) {
                return Array.Empty<ExcelColumnSnapshot>();
            }

            var result = new List<ExcelColumnSnapshot>();
            foreach (Column column in columns.Elements<Column>()) {
                int start = checked((int)(column.Min?.Value ?? 0U));
                int end = checked((int)(column.Max?.Value ?? 0U));
                if (start <= 0 || end <= 0 || end < start) {
                    continue;
                }

                result.Add(new ExcelColumnSnapshot {
                    StartIndex = start,
                    EndIndex = end,
                    Width = column.Width?.Value,
                    Hidden = column.Hidden?.Value == true,
                    CustomWidth = column.CustomWidth?.Value == true
                });
            }

            return result.AsReadOnly();
        }
    }
}
