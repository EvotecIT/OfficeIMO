using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table column.
    /// </summary>
    public sealed class PowerPointTableColumn {
        private readonly PowerPointTable _table;

        internal PowerPointTableColumn(PowerPointTable table, A.GridColumn column) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
            Column = column ?? throw new ArgumentNullException(nameof(column));
        }

        internal A.GridColumn Column { get; }

        /// <summary>
        ///     Gets the zero-based index of the column within the table.
        /// </summary>
        public int Index {
            get {
                int index = 0;
                foreach (A.GridColumn column in _table.TableElement.TableGrid?.Elements<A.GridColumn>() ?? Enumerable.Empty<A.GridColumn>()) {
                    if (ReferenceEquals(column, Column)) {
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
            get => Column.Width?.Value is { } extent ? OfficeLength.Emus(extent) : null;
            set => Column.Width = value?.ToEmus();
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

                return _table.TableElement.Elements<A.TableRow>()
                    .Select(row => _table.WrapCell(row.Elements<A.TableCell>().ElementAt(index)))
                    .ToList();
            }
        }

        /// <summary>
        ///     Retrieves a cell at the specified row index.
        /// </summary>
        public PowerPointTableCell GetCell(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= _table.Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }

            int index = Index;
            if (index < 0) {
                throw new InvalidOperationException("Column is no longer attached to the table.");
            }

            A.TableRow row = _table.TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            A.TableCell cell = row.Elements<A.TableCell>().ElementAt(index);
            return _table.WrapCell(cell);
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