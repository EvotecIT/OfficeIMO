using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table row.
    /// </summary>
    public sealed class PowerPointTableRow {
        private readonly PowerPointTable _table;
        private readonly PowerPointHeldNode _row;

        internal PowerPointTableRow(PowerPointTable table, OfficeXmlNode row) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
            if (row.IsNone) throw new ArgumentNullException(nameof(row));
            _row = new PowerPointHeldNode(row, table);
        }

        internal OfficeXmlNode Row => _row.Element;

        /// <summary>
        ///     Gets the zero-based index of the row within the table.
        /// </summary>
        public int Index {
            get {
                OfficeXmlNode current = Row;
                int index = 0;
                foreach (OfficeXmlNode row in _table.TableNode.Children(OfficeXmlNamespaces.DrawingML, "tr")) {
                    if (row.Equals(current)) {
                        return index;
                    }
                    index++;
                }
                return -1;
            }
        }

        /// <summary>
        ///     Gets or sets the row height as a physical length.
        /// </summary>
        public OfficeLength? Height {
            get => Row.Int64Attribute("h") is { } extent ? OfficeLength.Emus(extent) : null;
            set {
                Row.SetInt64Attribute("h", value?.ToEmus());
                _table.OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Row cells.
        /// </summary>
        public IReadOnlyList<PowerPointTableCell> Cells {
            get {
                var cells = new List<PowerPointTableCell>();
                foreach (OfficeXmlNode cell in Row.Children(OfficeXmlNamespaces.DrawingML, "tc")) cells.Add(_table.CellFor(cell));
                return cells;
            }
        }

        /// <summary>
        ///     Retrieves a cell at the specified column index.
        /// </summary>
        public PowerPointTableCell GetCell(int columnIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, _table.ColumnCount);

            int index = columnIndex;
            foreach (OfficeXmlNode cell in Row.Children(OfficeXmlNamespaces.DrawingML, "tc")) {
                if (index-- == 0) return _table.CellFor(cell);
            }
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        }

        /// <summary>
        ///     Removes this row from the table.
        /// </summary>
        public void Remove() {
            _table.RemoveRow(Row);
        }
    }
}
