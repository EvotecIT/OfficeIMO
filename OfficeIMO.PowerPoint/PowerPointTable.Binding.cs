using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Binds data to the table, expanding rows/columns as needed.
        /// </summary>
        public void Bind<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            bool includeHeaders = true, int startRow = 0, int startColumn = 0) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }
            if (columns == null) {
                throw new ArgumentNullException(nameof(columns));
            }
            if (startRow < 0) {
                throw new ArgumentOutOfRangeException(nameof(startRow));
            }
            if (startColumn < 0) {
                throw new ArgumentOutOfRangeException(nameof(startColumn));
            }

            var items = data.ToList();
            var columnList = columns.ToList();
            if (columnList.Count == 0) {
                throw new ArgumentException("At least one column is required.", nameof(columns));
            }

            int requiredRows = items.Count + (includeHeaders ? 1 : 0);
            int requiredColumns = columnList.Count;

            while (Rows < startRow + requiredRows) {
                AddRow();
            }
            while (Columns < startColumn + requiredColumns) {
                AddColumn();
            }

            int rowIndex = startRow;
            if (includeHeaders) {
                for (int c = 0; c < columnList.Count; c++) {
                    GetCell(rowIndex, startColumn + c).Text = columnList[c].Header;
                }
                rowIndex++;
            }

            foreach (var item in items) {
                for (int c = 0; c < columnList.Count; c++) {
                    object? value = columnList[c].ValueSelector(item);
                    string text = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    GetCell(rowIndex, startColumn + c).Text = text;
                }
                rowIndex++;
            }
        }
    }
}
