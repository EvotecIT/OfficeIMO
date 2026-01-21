namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Returns all pivot tables defined in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelPivotTableInfo> GetPivotTables() {
            return Locking.ExecuteRead(EnsureLock(), () => {
                using var _ = Locking.EnterNoLockScope();
                var list = new List<ExcelPivotTableInfo>();
                foreach (var sheet in Sheets) {
                    list.AddRange(sheet.GetPivotTables());
                }
                return list;
            });
        }
    }
}
