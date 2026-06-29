namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents BIFF8 enhanced worksheet-protection permissions from a FeatHdr ISFPROTECTION record.
    /// </summary>
    public sealed class LegacyXlsWorksheetProtectionPermissions : IEquatable<LegacyXlsWorksheetProtectionPermissions> {
        /// <summary>
        /// Creates worksheet-protection permission metadata.
        /// </summary>
        public LegacyXlsWorksheetProtectionPermissions(
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

        /// <summary>Gets whether editing worksheet objects is allowed.</summary>
        public bool AllowEditObjects { get; }

        /// <summary>Gets whether editing worksheet scenarios is allowed.</summary>
        public bool AllowEditScenarios { get; }

        /// <summary>Gets whether formatting cells is allowed.</summary>
        public bool AllowFormatCells { get; }

        /// <summary>Gets whether formatting columns is allowed.</summary>
        public bool AllowFormatColumns { get; }

        /// <summary>Gets whether formatting rows is allowed.</summary>
        public bool AllowFormatRows { get; }

        /// <summary>Gets whether inserting columns is allowed.</summary>
        public bool AllowInsertColumns { get; }

        /// <summary>Gets whether inserting rows is allowed.</summary>
        public bool AllowInsertRows { get; }

        /// <summary>Gets whether inserting hyperlinks is allowed.</summary>
        public bool AllowInsertHyperlinks { get; }

        /// <summary>Gets whether deleting columns is allowed.</summary>
        public bool AllowDeleteColumns { get; }

        /// <summary>Gets whether deleting rows is allowed.</summary>
        public bool AllowDeleteRows { get; }

        /// <summary>Gets whether selecting locked cells is allowed.</summary>
        public bool AllowSelectLockedCells { get; }

        /// <summary>Gets whether sorting is allowed.</summary>
        public bool AllowSort { get; }

        /// <summary>Gets whether using AutoFilter is allowed.</summary>
        public bool AllowAutoFilter { get; }

        /// <summary>Gets whether using PivotTables is allowed.</summary>
        public bool AllowPivotTables { get; }

        /// <summary>Gets whether selecting unlocked cells is allowed.</summary>
        public bool AllowSelectUnlockedCells { get; }

        internal static LegacyXlsWorksheetProtectionPermissions Default(bool? protectObjects, bool? protectScenarios) {
            return new LegacyXlsWorksheetProtectionPermissions(
                allowEditObjects: protectObjects != true,
                allowEditScenarios: protectScenarios != true,
                allowFormatCells: false,
                allowFormatColumns: false,
                allowFormatRows: false,
                allowInsertColumns: false,
                allowInsertRows: false,
                allowInsertHyperlinks: false,
                allowDeleteColumns: false,
                allowDeleteRows: false,
                allowSelectLockedCells: true,
                allowSort: false,
                allowAutoFilter: false,
                allowPivotTables: false,
                allowSelectUnlockedCells: true);
        }

        /// <inheritdoc />
        public bool Equals(LegacyXlsWorksheetProtectionPermissions? other) {
            return other != null
                && AllowEditObjects == other.AllowEditObjects
                && AllowEditScenarios == other.AllowEditScenarios
                && AllowFormatCells == other.AllowFormatCells
                && AllowFormatColumns == other.AllowFormatColumns
                && AllowFormatRows == other.AllowFormatRows
                && AllowInsertColumns == other.AllowInsertColumns
                && AllowInsertRows == other.AllowInsertRows
                && AllowInsertHyperlinks == other.AllowInsertHyperlinks
                && AllowDeleteColumns == other.AllowDeleteColumns
                && AllowDeleteRows == other.AllowDeleteRows
                && AllowSelectLockedCells == other.AllowSelectLockedCells
                && AllowSort == other.AllowSort
                && AllowAutoFilter == other.AllowAutoFilter
                && AllowPivotTables == other.AllowPivotTables
                && AllowSelectUnlockedCells == other.AllowSelectUnlockedCells;
        }

        /// <inheritdoc />
        public override bool Equals(object? obj) {
            return Equals(obj as LegacyXlsWorksheetProtectionPermissions);
        }

        /// <inheritdoc />
        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = (hash * 31) + AllowEditObjects.GetHashCode();
                hash = (hash * 31) + AllowEditScenarios.GetHashCode();
                hash = (hash * 31) + AllowFormatCells.GetHashCode();
                hash = (hash * 31) + AllowFormatColumns.GetHashCode();
                hash = (hash * 31) + AllowFormatRows.GetHashCode();
                hash = (hash * 31) + AllowInsertColumns.GetHashCode();
                hash = (hash * 31) + AllowInsertRows.GetHashCode();
                hash = (hash * 31) + AllowInsertHyperlinks.GetHashCode();
                hash = (hash * 31) + AllowDeleteColumns.GetHashCode();
                hash = (hash * 31) + AllowDeleteRows.GetHashCode();
                hash = (hash * 31) + AllowSelectLockedCells.GetHashCode();
                hash = (hash * 31) + AllowSort.GetHashCode();
                hash = (hash * 31) + AllowAutoFilter.GetHashCode();
                hash = (hash * 31) + AllowPivotTables.GetHashCode();
                hash = (hash * 31) + AllowSelectUnlockedCells.GetHashCode();
                return hash;
            }
        }
    }
}
