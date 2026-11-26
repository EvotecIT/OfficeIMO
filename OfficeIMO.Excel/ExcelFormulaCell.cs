namespace OfficeIMO.Excel {
    /// <summary>How a cell's formula is stored.</summary>
    internal enum ExcelFormulaKind : byte {
        Normal,
        Shared,
        Array,
        DataTable,
    }

    /// <summary>The outcome of reading a value from a sheet's saved pivot tables.</summary>
    internal enum ExcelPivotDataLookup : byte {
        Found,
        NotFound,
        Unsupported,
    }

    /// <summary>
    /// A worksheet cell as the formula engine reads and writes it. The handle carries the element of whichever model
    /// holds the sheet; an empty handle stands for a cell that isn't stored.
    /// </summary>
    internal readonly partial struct ExcelFormulaCell : IEquatable<ExcelFormulaCell> {
        private readonly object? _element;

        internal ExcelFormulaCell(object? element) => _element = element;

        internal bool IsEmpty => _element == null;

        internal object? Element => _element;

        public bool Equals(ExcelFormulaCell other) => ReferenceEquals(_element, other._element);

        public override bool Equals(object? obj) => obj is ExcelFormulaCell other && Equals(other);

        public override int GetHashCode() => _element == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_element);
    }
}

namespace OfficeIMO.Excel {
    /// <summary>A workbook sheet entry as formulas see it: its name, whether it is a worksheet, and its element.</summary>
    internal sealed class ExcelFormulaSheetEntry {
        internal ExcelFormulaSheetEntry(string? name, bool isWorksheet, object element) {
            Name = name;
            IsWorksheet = isWorksheet;
            Element = element;
        }

        internal string? Name { get; }
        internal bool IsWorksheet { get; }
        internal object Element { get; }
    }

    /// <summary>A defined name as formulas resolve it.</summary>
    internal sealed class ExcelFormulaDefinedName {
        internal ExcelFormulaDefinedName(string? name, string? text, uint? localSheetId) {
            Name = name;
            Text = text;
            LocalSheetId = localSheetId;
        }

        internal string? Name { get; }
        internal string? Text { get; }
        internal uint? LocalSheetId { get; }
    }

    /// <summary>A table as structured references resolve it.</summary>
    internal sealed class ExcelFormulaTable {
        internal ExcelFormulaTable(string? name, string? displayName, string? reference, uint headerRowCount,
            uint totalsRowCount, bool totalsRowShown, IReadOnlyList<string?> columnNames) {
            Name = name;
            DisplayName = displayName;
            Reference = reference;
            HeaderRowCount = headerRowCount;
            TotalsRowCount = totalsRowCount;
            TotalsRowShown = totalsRowShown;
            ColumnNames = columnNames;
        }

        internal string? Name { get; }
        internal string? DisplayName { get; }
        internal string? Reference { get; }
        internal uint HeaderRowCount { get; }
        internal uint TotalsRowCount { get; }
        internal bool TotalsRowShown { get; }
        internal IReadOnlyList<string?> ColumnNames { get; }
    }
}
