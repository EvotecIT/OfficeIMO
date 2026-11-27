using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Returns all Excel tables defined in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelTableInfo> GetTables() =>
            TryGetDeferredTables(null, out List<ExcelTableInfo> deferred) ? deferred : GetTables(null);

        /// <summary>
        /// Describes the table of a pending direct table on <paramref name="sheet"/> as loading its rows would leave it,
        /// so reading table metadata does not load the rows into the worksheet.
        /// </summary>
        internal bool TryGetDirectTabularSaveCandidateTable(ExcelSheet sheet, WorksheetPart worksheetPart, out ExcelTableInfo? table) {
            table = null;
            var candidate = _directDataSetSaveCandidate;
            if (!ReferenceEquals(sheet.Document, this)
                || candidate == null
                || !candidate.IsValid
                || !candidate.IsDeferred
                || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!sheetModel.HasTable
                || string.IsNullOrWhiteSpace(sheetModel.TableName)
                || string.IsNullOrWhiteSpace(sheetModel.Range)
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                return false;
            }

            // The worksheet may already hold this table's part; any other table needs the full listing.
            foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                if (!string.Equals(tablePart.Table?.Name?.Value, sheetModel.TableName, StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
            }

            int sheetIndex = -1;
            var sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            for (int i = 0; i < sheets.Count; i++) {
                if (string.Equals(sheets[i].Name?.Value, sheetModel.SheetName, StringComparison.OrdinalIgnoreCase)) {
                    sheetIndex = i;
                    break;
                }
            }

            var columns = new List<ExcelTableColumnInfo>(sheetModel.Table.ColumnCount);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < sheetModel.Table.ColumnCount; i++) {
                string name = sheetModel.Table.GetColumnName(i);
                if (string.IsNullOrWhiteSpace(name) || !names.Add(name)) {
                    return false;
                }

                columns.Add(new ExcelTableColumnInfo(i + 1, name));
            }

            table = new ExcelTableInfo(
                sheetModel.TableName!,
                sheetModel.TableName!,
                sheetModel.Range,
                sheetModel.SheetName,
                sheetIndex,
                sheetModel.TableStyle == ExcelTableStyle.None ? null : sheetModel.TableStyle.ToString(),
                sheetModel.IncludeHeaders,
                totalsRowShown: false,
                hasAutoFilter: sheetModel.IncludeAutoFilter && sheetModel.IncludeHeaders,
                sheetModel.ShowFirstColumn,
                sheetModel.ShowLastColumn,
                sheetModel.ShowRowStripes,
                sheetModel.ShowColumnStripes,
                columns);
            return true;
        }

        internal IReadOnlyList<ExcelTableInfo> GetTables(WorksheetPart? selectedWorksheet) {
            return ExecuteReadAfterMaterializing(() => {
                var result = new List<ExcelTableInfo>();
                var workbookPart = _spreadSheetDocument?.WorkbookPart;
                if (workbookPart == null) {
                    return result;
                }

                if (selectedWorksheet != null
                    && !workbookPart.WorksheetParts.Any(part => ReferenceEquals(part, selectedWorksheet))) {
                    throw new InvalidOperationException(
                        "The worksheet is no longer part of the current workbook package. Obtain a current worksheet from ExcelDocument.Sheets before inspecting or exporting it.");
                }

                var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
                var sheets = workbook.Sheets?.OfType<Sheet>().ToList() ?? new List<Sheet>();
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
                    if (selectedWorksheet != null && !ReferenceEquals(selectedWorksheet, worksheetPart)) {
                        continue;
                    }
                    var relId = workbookPart.GetIdOfPart(worksheetPart);
                    if (string.IsNullOrWhiteSpace(relId)) {
                        continue;
                    }

                    sheetLookup.TryGetValue(relId, out var sheetInfo);
                    var sheetName = sheetInfo.Name ?? string.Empty;
                    var sheetIndex = string.IsNullOrEmpty(sheetInfo.Name) ? -1 : sheetInfo.Index;

                    foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                        var table = tablePart.Table;
                        if (table == null) {
                            continue;
                        }

                        var name = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                        var displayName = table.DisplayName?.Value ?? name;
                        var range = table.Reference?.Value ?? string.Empty;
                        TableStyleInfo? styleInfo = table.TableStyleInfo;
                        var columns = table.TableColumns?.Elements<TableColumn>()
                            .Select((column, index) => new ExcelTableColumnInfo(
                                index + 1,
                                column.Name?.Value ?? string.Empty,
                                GetOpenXmlAttributeValue(column, "totalsRowFunction")))
                            .ToList() ?? new List<ExcelTableColumnInfo>();

                        result.Add(new ExcelTableInfo(
                            name,
                            displayName,
                            range,
                            sheetName,
                            sheetIndex,
                            styleInfo?.Name?.Value,
                            (table.HeaderRowCount?.Value ?? 1U) > 0U,
                            (table.TotalsRowCount?.Value ?? 0U) > 0U,
                            table.GetFirstChild<AutoFilter>() != null,
                            styleInfo?.ShowFirstColumn?.Value == true,
                            styleInfo?.ShowLastColumn?.Value == true,
                            styleInfo?.ShowRowStripes?.Value == true,
                            styleInfo?.ShowColumnStripes?.Value == true,
                            columns));
                    }
                }

                return result;
            });
        }
    }
}
