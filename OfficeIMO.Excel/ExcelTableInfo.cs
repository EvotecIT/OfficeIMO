using System.Collections.Generic;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a column inside an Excel table.
    /// </summary>
    public sealed class ExcelTableColumnInfo {
        /// <summary>Initializes a new instance of the <see cref="ExcelTableColumnInfo"/> class.</summary>
        /// <param name="index">1-based column index inside the table.</param>
        /// <param name="name">Column name.</param>
        /// <param name="totalsRowFunction">Totals row function name, when one is assigned.</param>
        public ExcelTableColumnInfo(int index, string name, string? totalsRowFunction = null) {
            Index = index;
            Name = name ?? string.Empty;
            TotalsRowFunction = totalsRowFunction;
        }

        /// <summary>1-based column index inside the table.</summary>
        public int Index { get; }

        /// <summary>Column name.</summary>
        public string Name { get; }

        /// <summary>Totals row function name, when one is assigned.</summary>
        public string? TotalsRowFunction { get; }
    }

    /// <summary>
    /// Represents a table defined in an Excel workbook.
    /// </summary>
    public sealed class ExcelTableInfo {
        /// <summary>Initializes a new instance of the <see cref="ExcelTableInfo"/> class.</summary>
        /// <param name="name">Table name (or display name).</param>
        /// <param name="range">Table range in A1 notation.</param>
        /// <param name="sheetName">Sheet name containing the table.</param>
        /// <param name="sheetIndex">0-based sheet index; -1 when unknown.</param>
        public ExcelTableInfo(string name, string range, string sheetName, int sheetIndex)
            : this(name, name, range, sheetName, sheetIndex, null, hasHeaderRow: true, totalsRowShown: false, hasAutoFilter: false, columns: null) {
        }

        internal ExcelTableInfo(
            string name,
            string displayName,
            string range,
            string sheetName,
            int sheetIndex,
            string? styleName,
            bool hasHeaderRow,
            bool totalsRowShown,
            bool hasAutoFilter,
            IReadOnlyList<ExcelTableColumnInfo>? columns) {
            Name = name ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
            Range = range ?? string.Empty;
            SheetName = sheetName ?? string.Empty;
            SheetIndex = sheetIndex;
            StyleName = styleName;
            HasHeaderRow = hasHeaderRow;
            TotalsRowShown = totalsRowShown;
            HasAutoFilter = hasAutoFilter;
            Columns = columns ?? System.Array.Empty<ExcelTableColumnInfo>();
        }

        /// <summary>Table name (or display name).</summary>
        public string Name { get; }

        /// <summary>Table display name.</summary>
        public string DisplayName { get; }

        /// <summary>Table range in A1 notation.</summary>
        public string Range { get; }

        /// <summary>Sheet name containing the table.</summary>
        public string SheetName { get; }

        /// <summary>0-based sheet index; -1 when unknown.</summary>
        public int SheetIndex { get; }

        /// <summary>Table style name, when present.</summary>
        public string? StyleName { get; }

        /// <summary>Whether the table has a header row.</summary>
        public bool HasHeaderRow { get; }

        /// <summary>Whether the table shows a totals row.</summary>
        public bool TotalsRowShown { get; }

        /// <summary>Whether the table has a table-scoped AutoFilter.</summary>
        public bool HasAutoFilter { get; }

        /// <summary>Table columns in display order.</summary>
        public IReadOnlyList<ExcelTableColumnInfo> Columns { get; }
    }
}
