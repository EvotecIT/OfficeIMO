using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets explicit worksheet row definitions such as custom heights and hidden rows.
        /// </summary>
        public IReadOnlyList<ExcelRowSnapshot> GetRowDefinitions() {
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return Array.Empty<ExcelRowSnapshot>();
            }

            var result = new List<ExcelRowSnapshot>();
            foreach (Row row in sheetData.Elements<Row>()) {
                int index = checked((int)(row.RowIndex?.Value ?? 0U));
                if (index <= 0) {
                    continue;
                }

                bool hidden = row.Hidden?.Value == true;
                bool customHeight = row.CustomHeight?.Value == true;
                if (!hidden && !customHeight && row.Height == null) {
                    continue;
                }

                result.Add(new ExcelRowSnapshot {
                    Index = index,
                    Height = row.Height?.Value,
                    Hidden = hidden,
                    CustomHeight = customHeight
                });
            }

            return result.AsReadOnly();
        }
    }
}
