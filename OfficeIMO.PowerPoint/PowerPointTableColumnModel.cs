using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table column.
    /// </summary>
    public sealed class PowerPointTableColumn {
        private const int D = OfficeXmlNamespaces.DrawingML;
        private readonly PowerPointTable _table;
        private readonly PowerPointHeldNode _column;

        internal PowerPointTableColumn(PowerPointTable table, OfficeXmlNode column) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
            if (column.IsNone) throw new ArgumentNullException(nameof(column));
            _column = new PowerPointHeldNode(column, table);
        }

        internal OfficeXmlNode Column => _column.Element;

        /// <summary>
        ///     Gets the zero-based index of the column within the table.
        /// </summary>
        public int Index {
            get {
                OfficeXmlNode current = Column;
                int index = 0;
                foreach (OfficeXmlNode column in _table.TableNode.Child(D, "tblGrid").Children(D, "gridCol")) {
                    if (column.Equals(current)) {
                        return index;
                    }
                    index++;
                }
                return -1;
            }
        }

        /// <summary>
        ///     Gets or sets the column width as a physical length.
        /// </summary>
        public OfficeLength? Width {
            get => Column.Int64Attribute("w") is { } extent ? OfficeLength.Emus(extent) : null;
            set {
                Column.SetInt64Attribute("w", value?.ToEmus());
                _table.OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Column cells.
        /// </summary>
        public IReadOnlyList<PowerPointTableCell> Cells {
            get {
                int index = Index;
                if (index < 0) {
                    return Array.Empty<PowerPointTableCell>();
                }

                var cells = new List<PowerPointTableCell>();
                foreach (OfficeXmlNode row in _table.TableNode.Children(D, "tr")) cells.Add(_table.CellFor(CellAt(row, index)));
                return cells;
            }
        }

        /// <summary>
        ///     Retrieves a cell at the specified row index.
        /// </summary>
        public PowerPointTableCell GetCell(int rowIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, _table.RowCount);

            int index = Index;
            if (index < 0) {
                throw new InvalidOperationException("Column is no longer attached to the table.");
            }

            int remaining = rowIndex;
            foreach (OfficeXmlNode row in _table.TableNode.Children(D, "tr")) {
                if (remaining-- == 0) return _table.CellFor(CellAt(row, index));
            }
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        }

        private static OfficeXmlNode CellAt(OfficeXmlNode row, int index) {
            foreach (OfficeXmlNode cell in row.Children(D, "tc")) {
                if (index-- == 0) return cell;
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        /// <summary>
        ///     Removes this column from the table.
        /// </summary>
        public void Remove() {
            int index = Index;
            if (index < 0) {
                return;
            }

            _table.RemoveColumn(index);
        }
    }
}
