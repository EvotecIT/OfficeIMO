using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>The <c>t</c> attribute of a stored cell; <see cref="Absent"/> leaves it out.</summary>
    internal enum ExcelStoredCellType : byte {
        Absent,
        SharedString,
        String,
        Number,
        Boolean,
        Error
    }

    /// <summary>A plain cell kept outside the DOM: its column, style, <c>t</c> attribute and value text.</summary>
    internal struct ExcelStoredCell {
        private const byte TypeMask = 0x07;
        private const byte HasStyleFlag = 0x08;
        private const byte HasValueFlag = 0x10;

        internal ExcelStoredCell(int column, uint style, bool hasStyle, ExcelStoredCellType type, string? value) {
            Column = column;
            Style = style;
            Kind = (byte)((byte)type | (hasStyle ? HasStyleFlag : 0) | (value != null ? HasValueFlag : 0));
            Value = value;
        }

        internal int Column;
        internal uint Style;
        internal byte Kind;
        internal string? Value;

        internal ExcelStoredCellType Type => (ExcelStoredCellType)(Kind & TypeMask);

        internal bool HasStyle => (Kind & HasStyleFlag) != 0;

        internal bool HasValue => (Kind & HasValueFlag) != 0;
    }

    /// <summary>
    /// A worksheet row that keeps its plain cells outside the SDK DOM. A plain cell has at most the <c>r</c>, <c>s</c>
    /// and <c>t</c> attributes and one value; formulas and everything else stay DOM children. A stored cell costs a
    /// few bytes plus its value text instead of a <see cref="Cell"/>, its value element and their attribute objects,
    /// which is most of the cost of loading and editing a worksheet.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Every SDK route to an element's children starts at <see cref="FirstChild"/>, <see cref="LastChild"/> or
    /// <see cref="HasChildren"/>: enumeration, descendant search, sibling links, serialization, cloning and
    /// validation. Those members, and every child mutation, first move the stored cells into the DOM, so no code can
    /// observe a row without them. Code that knows the store reads both inside <see cref="BeginRawAccess"/>.
    /// </para>
    /// <para>
    /// A cell is either stored or a DOM child, never both. Stored cells are in ascending column order, and every DOM
    /// cell of a row with stored cells has a canonical reference on the row, so writers merge the two by column.
    /// Stored cells take their reference from the row index, so moving a row moves them as it moves child cells.
    /// </para>
    /// </remarks>
    internal sealed class ExcelStoredRow : Row {
        [ThreadStatic]
        private static int _rawAccessDepth;

        private static readonly string[] TypeNames = { string.Empty, "s", "str", "n", "b", "e" };

        private readonly int _ownerThreadId;
        private ExcelStoredCell[]? _cells;
        private int _count;

        internal ExcelStoredRow(int ownerThreadId) => _ownerThreadId = ownerThreadId;

        /// <summary>Gets the stored cells, or null once they have moved to the DOM.</summary>
        internal ExcelStoredCell[]? Cells => _cells;

        internal int StoredCount => _cells == null ? 0 : _count;

        /// <summary>
        /// Gets whether the calling code may read and change the stored cells. Another thread sees only the DOM, which
        /// materializes the row on first access.
        /// </summary>
        internal bool IsStoreAvailable => _cells != null && Environment.CurrentManagedThreadId == _ownerThreadId;

        /// <summary>
        /// Gets whether a new plain cell may be stored in this row. A DOM cell without a canonical reference, or one
        /// out of column order, could not be merged with stored cells by column, so such a row keeps its cells in the DOM.
        /// </summary>
        internal bool AcceptsStoredCells => !_hasUnorderedDomCells && Environment.CurrentManagedThreadId == _ownerThreadId;

        private bool _hasUnorderedDomCells;

        /// <summary>Gets whether every DOM cell has a canonical reference in column order with the stored cells.</summary>
        internal bool HasOrderedCells => !_hasUnorderedDomCells;

        /// <summary>Records that the loader kept a DOM cell it could not place by column.</summary>
        internal void MarkUnorderedDomCells() => _hasUnorderedDomCells = true;

        internal static string TypeName(ExcelStoredCellType type) => TypeNames[(int)type];

        internal static bool TryGetType(string? name, out ExcelStoredCellType type) {
            for (int index = 1; index < TypeNames.Length; index++) {
                if (string.Equals(TypeNames[index], name, StringComparison.Ordinal)) {
                    type = (ExcelStoredCellType)index;
                    return true;
                }
            }
            type = ExcelStoredCellType.Absent;
            return name == null;
        }

        /// <summary>Lets the caller walk this row's DOM children without moving the stored cells into the DOM.</summary>
        internal static RawAccessScope BeginRawAccess() => BeginRawAccessIf(true);

        /// <summary>Begins raw access when <paramref name="active"/> is true; otherwise the scope does nothing.</summary>
        internal static RawAccessScope BeginRawAccessIf(bool active) {
            if (active) _rawAccessDepth++;
            return new RawAccessScope(active);
        }

        internal readonly struct RawAccessScope : IDisposable {
            private readonly bool _active;

            internal RawAccessScope(bool active) => _active = active;

            public void Dispose() {
                if (_active) _rawAccessDepth--;
            }
        }

        private bool MustMaterialize => _cells != null && _rawAccessDepth == 0;

        /// <summary>Takes the loader's cells for this row; the array is owned by the row from now on.</summary>
        internal void SetStoredCells(ExcelStoredCell[] cells, int count) {
            _cells = count == 0 ? null : cells;
            _count = count;
        }

        internal ref ExcelStoredCell CellAt(int index) => ref _cells![index];

        /// <summary>Returns the index of <paramref name="column"/>, or the bitwise complement of where it would go.</summary>
        internal int Find(int column) {
            ExcelStoredCell[]? cells = _cells;
            if (cells == null) return ~0;
            int count = _count;
            // Rows are written and read left to right, so the last cell answers most lookups.
            if (count > 0 && cells[count - 1].Column <= column) {
                return cells[count - 1].Column == column ? count - 1 : ~count;
            }
            int low = 0, high = count - 1;
            while (low <= high) {
                int middle = low + ((high - low) >> 1);
                int candidate = cells[middle].Column;
                if (candidate == column) return middle;
                if (candidate < column) low = middle + 1; else high = middle - 1;
            }
            return ~low;
        }

        /// <summary>Stores a new cell at the position <see cref="Find"/> reported. The column must not be a DOM cell.</summary>
        internal void Insert(int index, in ExcelStoredCell cell) {
            ExcelStoredCell[] cells = _cells ?? new ExcelStoredCell[4];
            if (_cells == null) _count = 0;
            if (_count == cells.Length) Array.Resize(ref cells, cells.Length * 2);
            if (index < _count) Array.Copy(cells, index, cells, index + 1, _count - index);
            cells[index] = cell;
            _cells = cells;
            _count++;
        }

        internal void RemoveAt(int index) {
            ExcelStoredCell[] cells = _cells!;
            int moved = _count - index - 1;
            if (moved > 0) Array.Copy(cells, index + 1, cells, index, moved);
            cells[--_count] = default;
            if (_count == 0) _cells = null;
        }

        /// <summary>Drops every stored cell, as removing them from the DOM would.</summary>
        internal void ClearStoredCells() {
            _cells = null;
            _count = 0;
        }

        /// <summary>Applies the column insertion or deletion of <paramref name="count"/> columns at <paramref name="firstColumn"/>.</summary>
        internal bool ShiftStoredColumns(int firstColumn, int count, bool deleting) {
            ExcelStoredCell[]? cells = _cells;
            if (cells == null) return false;
            bool changed = false;
            int lastColumn = firstColumn + count - 1;
            for (int index = 0; index < _count; index++) {
                int column = cells[index].Column;
                if (deleting) {
                    if (column >= firstColumn && column <= lastColumn) {
                        RemoveAt(index--);
                        changed = true;
                        if (_cells == null) return true;
                    } else if (column > lastColumn) {
                        cells[index].Column = column - count;
                        changed = true;
                    }
                } else if (column >= firstColumn) {
                    cells[index].Column = checked(column + count);
                    changed = true;
                }
            }
            return changed;
        }

        /// <summary>
        /// Builds the DOM form of a stored cell: the attributes and value the grid loader would have given it, so the
        /// cell serializes to the same bytes.
        /// </summary>
        internal static Cell CreateCell(string reference, in ExcelStoredCell stored) {
            var cell = new Cell { CellReference = reference };
            if (stored.HasStyle) cell.StyleIndex = new UInt32Value { InnerText = InvariantNumberText.Get(stored.Style) };
            ExcelStoredCellType type = stored.Type;
            if (type != ExcelStoredCellType.Absent) cell.DataType = new EnumValue<CellValues> { InnerText = TypeName(type) };
            if (stored.HasValue) cell.AppendChild(new CellValue(stored.Value ?? string.Empty));
            return cell;
        }

        /// <summary>Moves one stored cell into the DOM at its column position and returns it.</summary>
        internal Cell Promote(int index) {
            ExcelStoredCell stored = _cells![index];
            RemoveAt(index);
            var cell = CreateCell(ColumnLetters(stored.Column) + InvariantNumberText.Get(checked((int)RowIndex!.Value)), stored);
            using (BeginRawAccess()) {
                InsertDomCell(cell, stored.Column);
            }
            return cell;
        }

        /// <summary>Moves every stored cell into the DOM, merged with the DOM cells by column.</summary>
        internal void Materialize() {
            if (_cells == null) return;
            lock (this) {
                ExcelStoredCell[]? cells = _cells;
                if (cells == null) return;
                int count = _count;
                using (BeginRawAccess()) {
                    int rowIndex = checked((int)RowIndex!.Value);
                    string rowText = InvariantNumberText.Get(rowIndex);
                    OpenXmlElement? cursor = base.FirstChild;
                    int cursorColumn = ColumnOf(cursor);
                    for (int index = 0; index < count; index++) {
                        ref ExcelStoredCell stored = ref cells[index];
                        while (cursor != null && cursorColumn < stored.Column) {
                            cursor = cursor.NextSibling();
                            cursorColumn = ColumnOf(cursor);
                        }
                        Cell cell = CreateCell(ColumnLetters(stored.Column) + rowText, stored);
                        if (cursor == null) base.AppendChild(cell); else base.InsertBefore(cell, cursor);
                    }
                }
                _cells = null;
                _count = 0;
            }
        }

        private static readonly string?[] ColumnLetterCache = new string?[ExcelA1.MaxColumns + 1];

        /// <summary>Shares column letters between materialized rows, so a cell's reference costs one string as when loaded.</summary>
        private static string ColumnLetters(int column) => ColumnLetterCache[column] ??= ExcelA1.ColumnIndexToLetters(column);

        private void InsertDomCell(Cell cell, int column) {
            for (OpenXmlElement? child = base.FirstChild; child != null; child = child.NextSibling()) {
                if (ColumnOf(child) > column) {
                    base.InsertBefore(cell, child);
                    return;
                }
            }
            base.AppendChild(cell);
        }

        /// <summary>Gets the column of a DOM cell; anything else sorts after every stored cell.</summary>
        private static int ColumnOf(OpenXmlElement? element) {
            if (element == null) return int.MaxValue;
            if (element is not Cell cell) return int.MaxValue;
            int column = ExcelA1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
            return column > 0 ? column : int.MaxValue;
        }

        private void MaterializeIfNeeded() {
            if (MustMaterialize) Materialize();
        }

        public override OpenXmlElement? FirstChild {
            get {
                MaterializeIfNeeded();
                return base.FirstChild;
            }
        }

        public override OpenXmlElement? LastChild {
            get {
                MaterializeIfNeeded();
                return base.LastChild;
            }
        }

        public override bool HasChildren {
            get {
                MaterializeIfNeeded();
                return base.HasChildren;
            }
        }

        public override string InnerText {
            get {
                MaterializeIfNeeded();
                return base.InnerText;
            }
        }

        public override T? AppendChild<T>(T? newChild) where T : class {
            MaterializeIfNeeded();
            return base.AppendChild(newChild);
        }

        public override T? InsertAfter<T>(T? newChild, OpenXmlElement? referenceChild) where T : class {
            MaterializeIfNeeded();
            return base.InsertAfter(newChild, referenceChild);
        }

        public override T? InsertBefore<T>(T? newChild, OpenXmlElement? referenceChild) where T : class {
            MaterializeIfNeeded();
            return base.InsertBefore(newChild, referenceChild);
        }

        public override T? InsertAt<T>(T? newChild, int index) where T : class {
            MaterializeIfNeeded();
            return base.InsertAt(newChild, index);
        }

        public override T? PrependChild<T>(T? newChild) where T : class {
            MaterializeIfNeeded();
            return base.PrependChild(newChild);
        }

        public override T? RemoveChild<T>(T? oldChild) where T : class {
            MaterializeIfNeeded();
            return base.RemoveChild(oldChild);
        }

        public override T? ReplaceChild<T>(OpenXmlElement newChild, T? oldChild) where T : class {
            MaterializeIfNeeded();
            return base.ReplaceChild(newChild, oldChild);
        }

        public override void RemoveAllChildren() {
            MaterializeIfNeeded();
            base.RemoveAllChildren();
        }

        public override OpenXmlElement CloneNode(bool deep) {
            MaterializeIfNeeded();
            return base.CloneNode(deep);
        }

        public override void WriteTo(System.Xml.XmlWriter xmlWriter) {
            MaterializeIfNeeded();
            base.WriteTo(xmlWriter);
        }
    }
}
