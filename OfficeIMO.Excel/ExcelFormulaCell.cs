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
    /// A stored worksheet cell as the formula engine reads and writes it. The worksheet XML the workbook loaded gives
    /// one kind of record; a worksheet held by another model gives its own.
    /// </summary>
    internal abstract class ExcelFormulaCellRecord {
        /// <summary>What two records of one cell share, so handles to the cell compare equal.</summary>
        internal abstract object Identity { get; }

        /// <summary>Whether the cell holds a formula element of its own.</summary>
        internal abstract bool HasFormula { get; }

        /// <summary>The formula text as stored; empty for a shared formula's dependent cells.</summary>
        internal abstract string? FormulaText { get; }

        /// <summary>How the formula is stored; <see cref="ExcelFormulaKind.Normal"/> when the cell has none.</summary>
        internal abstract ExcelFormulaKind FormulaKind { get; }

        /// <summary>The shared formula group the formula belongs to.</summary>
        internal abstract uint? SharedIndex { get; }

        /// <summary>The range the formula covers (array and shared formula masters).</summary>
        internal abstract string? FormulaReference { get; }

        /// <summary>Whether the formula asks to be calculated on the next pass; null when the flag is absent or unreadable.</summary>
        internal abstract bool? CalculateCell { get; }

        /// <summary>The A1 reference the cell records.</summary>
        internal abstract string? Reference { get; }

        /// <summary>The stored value text, or null when the cell has no value element.</summary>
        internal abstract string? ValueText { get; }

        /// <summary>The stored type of the value; null when the cell names none (a number).</summary>
        internal abstract ExcelCellValueType? ValueType { get; }

        /// <summary>The value metadata index (rich values and dynamic array results).</summary>
        internal abstract uint? ValueMetaIndex { get; }

        /// <summary>The cell metadata index (dynamic array formulas).</summary>
        internal abstract uint? CellMetaIndex { get; }

        /// <summary>Whether the cell names a cell metadata index, read or not.</summary>
        internal abstract bool HasCellMetaIndex { get; }

        /// <summary>The cell's style index.</summary>
        internal abstract uint? StyleIndex { get; }

        /// <summary>Whether the cell holds an inline string.</summary>
        internal abstract bool HasInlineString { get; }

        /// <summary>The text of the cell's inline string, decoded; empty when it has none.</summary>
        internal abstract string InlineText { get; }

        /// <summary>Whether the cell has any child element.</summary>
        internal abstract bool HasChildren { get; }

        /// <summary>The cell's effective position, which a cell without a reference takes from its place in its row.</summary>
        internal abstract bool TryGetPosition(out int row, out int column);

        /// <summary>The error a rich value names for an error cell, or <paramref name="fallback"/>.</summary>
        internal abstract string? ResolveRichValueError(string? fallback);

        /// <summary>Sets or removes (null) the formula's calculate-on-load flag, leaving an equal value as it is.</summary>
        internal abstract void SetCalculateCell(bool? calculate);

        /// <summary>Sets the value text (null removes the value element) and the value type (null leaves the type out).</summary>
        internal abstract void SetValue(string? text, ExcelCellValueType? type);

        /// <summary>Removes the cell's inline string.</summary>
        internal abstract void ClearInlineString();

        /// <summary>Removes the cell's value metadata index.</summary>
        internal abstract void ClearValueMetaIndex();

        /// <summary>Sets the range the formula covers.</summary>
        internal abstract void SetFormulaReference(string reference);

        /// <summary>Gives a shared formula's member its own formula text and calculates it on the next pass.</summary>
        internal abstract void SetOwnFormula(string text);
    }

    /// <summary>
    /// A worksheet cell as the formula engine reads and writes it. The handle carries the record of whichever model
    /// holds the sheet; an empty handle stands for a cell that isn't stored.
    /// </summary>
    internal readonly struct ExcelFormulaCell : IEquatable<ExcelFormulaCell> {
        private readonly ExcelFormulaCellRecord? _record;

        internal ExcelFormulaCell(ExcelFormulaCellRecord? record) => _record = record;

        internal bool IsEmpty => _record == null;

        internal ExcelFormulaCellRecord? Record => _record;

        internal bool HasFormula => _record?.HasFormula == true;

        internal string? FormulaText => _record?.FormulaText;

        internal ExcelFormulaKind FormulaKind => _record?.FormulaKind ?? ExcelFormulaKind.Normal;

        internal uint? SharedIndex => _record?.SharedIndex;

        internal string? FormulaReference => _record?.FormulaReference;

        internal bool? CalculateCell => _record?.CalculateCell;

        internal string? Reference => _record?.Reference;

        internal string? ValueText => _record?.ValueText;

        internal bool HasValue => _record?.ValueText != null;

        internal ExcelCellValueType? ValueType => _record?.ValueType;

        internal uint? ValueMetaIndex => _record?.ValueMetaIndex;

        internal uint? CellMetaIndex => _record?.CellMetaIndex;

        internal bool HasCellMetaIndex => _record?.HasCellMetaIndex == true;

        internal uint? StyleIndex => _record?.StyleIndex;

        internal bool HasInlineString => _record?.HasInlineString == true;

        /// <summary>Whether the cell holds anything a spilled value would cover: a value, text, a formula or metadata.</summary>
        internal bool HasContent => _record is { } record && (record.ValueText != null || record.HasInlineString
            || record.HasFormula || record.HasCellMetaIndex || record.ValueMetaIndex != null);

        /// <summary>Sets whether the formula is calculated on the next pass, leaving an equal value as it is.</summary>
        internal void SetCalculateCell(bool calculate) => _record!.SetCalculateCell(calculate);

        /// <summary>Removes the formula's calculate-on-load flag.</summary>
        internal void ClearCalculateCell() => _record!.SetCalculateCell(null);

        public bool Equals(ExcelFormulaCell other) =>
            ReferenceEquals(_record, other._record) || (_record != null && other._record != null && ReferenceEquals(_record.Identity, other._record.Identity));

        public override bool Equals(object? obj) => obj is ExcelFormulaCell other && Equals(other);

        public override int GetHashCode() => _record == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_record.Identity);
    }

    /// <summary>
    /// A worksheet as the formula engine sees it: its cells, merged ranges and tables. The workbook's own XML backs it
    /// while the workbook has not opened another model.
    /// </summary>
    internal abstract class ExcelFormulaSheetModel {
        /// <summary>The worksheet part, the same for every wrapper of one worksheet.</summary>
        internal abstract Uri PartUri { get; }

        /// <summary>The stored cell at the position, or null.</summary>
        internal abstract ExcelFormulaCellRecord? Find(int row, int column);

        /// <summary>The stored cell at the position, created empty when there is none.</summary>
        internal abstract ExcelFormulaCellRecord GetOrCreate(int row, int column);

        /// <summary>Removes a cell that has no style and no content left.</summary>
        internal abstract void RemoveIfEmpty(ExcelFormulaCellRecord cell);

        /// <summary>The cells that hold a formula element, in document order.</summary>
        internal abstract IEnumerable<ExcelFormulaCellRecord> FormulaCells();

        /// <summary>Every stored cell, in document order.</summary>
        internal abstract IEnumerable<ExcelFormulaCellRecord> Cells();

        /// <summary>The effective coordinates of every cell, including cells that omit their reference.</summary>
        internal abstract IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> EffectiveCoordinates();

        /// <summary>The references of the sheet's merged ranges.</summary>
        internal abstract IEnumerable<string> MergedRanges();

        /// <summary>The sheet's tables.</summary>
        internal abstract IReadOnlyList<ExcelFormulaTable> Tables();

        /// <summary>Whether the sheet has tables, whose header rows hold literal column names.</summary>
        internal abstract bool HasTableHeaders { get; }

        /// <summary>Whether the range covers a cell of a table's header row.</summary>
        internal abstract bool IsTableHeader(int top, int left, int bottom, int right);

        /// <summary>The cell's value as cell reads give it.</summary>
        internal abstract ExcelCellData ReadValue(int row, int column);

        /// <summary>Reads a plain number or shared-string cell without building a value snapshot.</summary>
        internal abstract bool TryReadPlainValue(int row, int column, out double? number, out string? text);

        /// <summary>Whether the sheet has saved pivot tables whose cells formulas may read.</summary>
        internal abstract bool HasPivotTables { get; }
    }

    /// <summary>A workbook as the formula engine sees it: its sheets, defined names, tables and cell metadata.</summary>
    internal abstract class ExcelFormulaWorkbookModel {
        /// <summary>The workbook's sheet entries in order, each marked as a worksheet or not.</summary>
        internal abstract List<ExcelFormulaSheetEntry> SheetEntries();

        /// <summary>The workbook's defined names in order.</summary>
        internal abstract List<ExcelFormulaDefinedName> DefinedNames();

        /// <summary>A sheet wrapper for a workbook sheet entry, outside the document's sheet cache.</summary>
        internal abstract ExcelSheet CreateSheet(ExcelFormulaSheetEntry entry);

        /// <summary>The tables of every worksheet with the sheet entry that holds each.</summary>
        internal abstract IEnumerable<(ExcelFormulaSheetEntry Sheet, ExcelFormulaTable Table)> Tables();

        /// <summary>Whether the workbook has a cell metadata part.</summary>
        internal abstract bool HasCellMetadata { get; }

        /// <summary>The workbook's cell metadata, or null when it has none.</summary>
        internal abstract ExcelCellMetadataIndex? CellMetadata { get; }

        /// <summary>Whether the workbook or the sheet asks for a full calculation when it opens.</summary>
        internal abstract bool RequestsFullCalculation(ExcelSheet sheet);
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
