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
        ///     Gets or sets the row height in EMUs.
        /// </summary>
        public long? HeightEmus {
            get => Row.Height?.Value;
            set => Row.Height = value;
        }

        /// <summary>
        ///     Gets or sets the row height in points.
        /// </summary>
        public double? HeightPoints {
            get => Row.Height?.Value != null ? PowerPointUnits.ToPoints(Row.Height.Value) : null;
            set => Row.Height = value != null ? PowerPointUnits.FromPoints(value.Value) : null;
        }

        /// <summary>
        ///     Gets or sets the row height in centimeters.
        /// </summary>
        public double? HeightCm {
            get => Row.Height?.Value != null ? PowerPointUnits.ToCentimeters(Row.Height.Value) : null;
            set => Row.Height = value != null ? PowerPointUnits.FromCentimeters(value.Value) : null;
        }

        /// <summary>
        ///     Gets or sets the row height in inches.
        /// </summary>
        public double? HeightInches {
            get => Row.Height?.Value != null ? PowerPointUnits.ToInches(Row.Height.Value) : null;
            set => Row.Height = value != null ? PowerPointUnits.FromInches(value.Value) : null;
        }

        /// <summary>
        ///     Row cells.
        /// </summary>
        public IReadOnlyList<PowerPointTableCell> Cells =>
            Row.Elements<A.TableCell>().Select(cell => new PowerPointTableCell(cell)).ToList();

        /// <summary>
        ///     Retrieves a cell at the specified column index.
        /// </summary>
        public PowerPointTableCell GetCell(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= _table.Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }

            A.TableCell cell = Row.Elements<A.TableCell>().ElementAt(columnIndex);
            return new PowerPointTableCell(cell);
        }

        /// <summary>
        ///     Removes this row from the table.
        /// </summary>
        public void Remove() {
            Row.Remove();
        }
    }
}
