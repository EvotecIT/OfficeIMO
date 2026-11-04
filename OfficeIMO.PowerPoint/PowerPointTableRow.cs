using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table row.
    /// </summary>
    public sealed class PowerPointTableRow {
        private readonly PowerPointTable _table;

        internal PowerPointTableRow(PowerPointTable table, A.TableRow row) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
            Row = row ?? throw new ArgumentNullException(nameof(row));
        }

        internal A.TableRow Row { get; }

        /// <summary>
        ///     Gets the zero-based index of the row within the table.
        /// </summary>
        public int Index {
            get {
                int index = 0;
                foreach (A.TableRow row in _table.TableElement.Elements<A.TableRow>()) {
                    if (ReferenceEquals(row, Row)) {
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
            get => Row.Height?.Value is { } extent ? OfficeLength.Emus(extent) : null;
            set => Row.Height = value?.ToEmus();
        }

        /// <summary>
        ///     Row cells.
        /// </summary>
        public IReadOnlyList<PowerPointTableCell> Cells =>
            Row.Elements<A.TableCell>().Select(cell =>
                _table.WrapCell(cell)).ToList();

        /// <summary>
        ///     Retrieves a cell at the specified column index.
        /// </summary>
        public PowerPointTableCell GetCell(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= _table.Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }

            A.TableCell cell = Row.Elements<A.TableCell>().ElementAt(columnIndex);
            return _table.WrapCell(cell);
        }

        /// <summary>
        ///     Removes this row from the table.
        /// </summary>
        public void Remove() {
            _table.RemoveRow(Row);
        }
    }
}