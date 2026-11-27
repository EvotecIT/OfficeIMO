namespace OfficeIMO.Excel {
    /// <summary>
    /// Table reads of a workbook that has not opened its SDK package. Recorded cell values never change a table (a
    /// value in a header row opens the SDK), so the tables are read from the package the workbook holds.
    /// </summary>
    public partial class ExcelDocument {
        /// <summary>
        /// Lists the tables of every worksheet, or of <paramref name="selected"/> only, as the model lists them: worksheets
        /// in the order the workbook part relates them, tables in the order each worksheet relates them. False when the
        /// workbook reads its tables through the SDK model.
        /// </summary>
        internal bool TryGetDeferredTables(ExcelDeferredSheet? selected, out List<ExcelTableInfo> tables) {
            tables = null!;
            if (TryGetXmlFormulaWorkbook() is not { } workbook) return false;
            List<ExcelTableInfo> result = new();
            bool read = Locking.ExecuteRead(EnsureLock(), () => {
                if (!ReferenceEquals(TryGetXmlFormulaWorkbook(), workbook)) return false;
                List<ExcelDeferredSheet> sheets = workbook.Workbook.Sheets;
                if (selected != null && !sheets.Contains(selected)) {
                    throw new InvalidOperationException(
                        "The worksheet is no longer part of the current workbook package. Obtain a current worksheet from ExcelDocument.Sheets before inspecting or exporting it.");
                }
                foreach (string relationshipId in workbook.WorksheetRelationshipOrder()) {
                    int index = sheets.FindIndex(sheet => string.Equals(sheet.RelationshipId, relationshipId, StringComparison.Ordinal));
                    if (index < 0 || !sheets[index].IsWorksheet) continue;
                    ExcelDeferredSheet sheet = sheets[index];
                    if (selected != null && !ReferenceEquals(sheet, selected)) continue;
                    foreach (ExcelTablePart table in workbook.ReadTableParts(sheet)) result.Add(table.ToInfo(sheet.Name, index));
                }
                return true;
            });
            if (!read) return false;
            tables = result;
            return true;
        }

        /// <summary>The table parts of a worksheet that has not opened its SDK package, or null when it reads them through the model.</summary>
        internal List<ExcelTablePart>? TryGetDeferredTableParts(ExcelDeferredSheet sheet) =>
            TryGetXmlFormulaWorkbook() is { } workbook && workbook.Workbook.Sheets.Contains(sheet) ? workbook.ReadTableParts(sheet) : null;
    }
}
