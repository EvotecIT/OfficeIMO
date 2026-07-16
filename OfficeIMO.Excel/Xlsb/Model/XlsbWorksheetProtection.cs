namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents classic BIFF12 worksheet protection and its allowed actions.</summary>
    internal sealed class XlsbWorksheetProtection {
        internal XlsbWorksheetProtection(
            ushort password,
            bool isProtected,
            bool allowEditObjects,
            bool allowEditScenarios,
            bool allowFormatCells,
            bool allowFormatColumns,
            bool allowFormatRows,
            bool allowInsertColumns,
            bool allowInsertRows,
            bool allowInsertHyperlinks,
            bool allowDeleteColumns,
            bool allowDeleteRows,
            bool allowSelectLockedCells,
            bool allowSort,
            bool allowAutoFilter,
            bool allowPivotTables,
            bool allowSelectUnlockedCells) {
            Password = password;
            IsProtected = isProtected;
            AllowEditObjects = allowEditObjects;
            AllowEditScenarios = allowEditScenarios;
            AllowFormatCells = allowFormatCells;
            AllowFormatColumns = allowFormatColumns;
            AllowFormatRows = allowFormatRows;
            AllowInsertColumns = allowInsertColumns;
            AllowInsertRows = allowInsertRows;
            AllowInsertHyperlinks = allowInsertHyperlinks;
            AllowDeleteColumns = allowDeleteColumns;
            AllowDeleteRows = allowDeleteRows;
            AllowSelectLockedCells = allowSelectLockedCells;
            AllowSort = allowSort;
            AllowAutoFilter = allowAutoFilter;
            AllowPivotTables = allowPivotTables;
            AllowSelectUnlockedCells = allowSelectUnlockedCells;
        }

        internal ushort Password { get; }
        internal bool IsProtected { get; }
        internal bool AllowEditObjects { get; }
        internal bool AllowEditScenarios { get; }
        internal bool AllowFormatCells { get; }
        internal bool AllowFormatColumns { get; }
        internal bool AllowFormatRows { get; }
        internal bool AllowInsertColumns { get; }
        internal bool AllowInsertRows { get; }
        internal bool AllowInsertHyperlinks { get; }
        internal bool AllowDeleteColumns { get; }
        internal bool AllowDeleteRows { get; }
        internal bool AllowSelectLockedCells { get; }
        internal bool AllowSort { get; }
        internal bool AllowAutoFilter { get; }
        internal bool AllowPivotTables { get; }
        internal bool AllowSelectUnlockedCells { get; }
    }
}
