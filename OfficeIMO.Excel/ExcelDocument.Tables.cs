using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Returns all Excel tables defined in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelTableInfo> GetTables() {
            return Locking.ExecuteRead(EnsureLock(), () => {
                var result = new List<ExcelTableInfo>();
                var workbookPart = _spreadSheetDocument?.WorkbookPart;
                if (workbookPart == null) {
                    return result;
                }

                var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>().ToList() ?? new List<Sheet>();
                var sheetLookup = new Dictionary<string, (string Name, int Index)>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < sheets.Count; i++) {
                    var sheet = sheets[i];
                    var id = sheet.Id?.Value;
                    var name = sheet.Name?.Value;
                    if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) {
                        continue;
                    }
                    sheetLookup[id!] = (name!, i);
                }

                foreach (var worksheetPart in workbookPart.WorksheetParts) {
                    var relId = workbookPart.GetIdOfPart(worksheetPart);
                    if (string.IsNullOrWhiteSpace(relId)) {
                        continue;
                    }

                    sheetLookup.TryGetValue(relId, out var sheetInfo);
                    var sheetName = sheetInfo.Name ?? string.Empty;
                    var sheetIndex = sheetInfo.Name == null ? -1 : sheetInfo.Index;

                    foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                        var table = tablePart.Table;
                        if (table == null) {
                            continue;
                        }

                        var name = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                        var range = table.Reference?.Value ?? string.Empty;
                        result.Add(new ExcelTableInfo(name, range, sheetName, sheetIndex));
                    }
                }

                return result;
            });
        }
    }
}
