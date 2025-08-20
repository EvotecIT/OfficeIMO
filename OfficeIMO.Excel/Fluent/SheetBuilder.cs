using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel.Utilities;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel.Fluent {
    public class SheetBuilder {
        private readonly ExcelFluentWorkbook _fluent;
        internal ExcelSheet? Sheet { get; private set; }
        private int _currentRow = 1;

        internal SheetBuilder(ExcelFluentWorkbook fluent) {
            _fluent = fluent;
        }

        public SheetBuilder AddSheet(string name = "") {
            Sheet = _fluent.Workbook.AddWorkSheet(name);
            return this;
        }

        public SheetBuilder HeaderRow(params object[] values) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var row = new RowBuilder(this, Sheet, _currentRow);
            row.Values(values);
            _currentRow++;
            return this;
        }

        public SheetBuilder Row(Action<RowBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new RowBuilder(this, Sheet, _currentRow);
            action(builder);
            _currentRow++;
            return this;
        }

        public SheetBuilder RowsFrom<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (data == null) throw new ArgumentNullException(nameof(data));

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            var flattener = new ObjectFlattener();

            List<string>? headers = null;
            bool headerWritten = false;
            foreach (var item in data) {
                var dict = flattener.Flatten(item, options);
                if (!headerWritten) {
                    headers = new List<string>(dict.Keys);
                    HeaderRow(headers.Cast<object>().ToArray());
                    headerWritten = true;
                }
                Row(r => r.Values(headers!.Select(h => dict.TryGetValue(h, out var v) ? v : null).ToArray()));
            }

            return this;
        }

        public SheetBuilder Cell(int row, int column, object value) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
            Sheet.SetCellValue(row, column, value);
            return this;
        }

        public SheetBuilder Range(int fromRow, int fromCol, int toRow, int toCol, object[,]? values = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (fromRow <= 0) throw new ArgumentOutOfRangeException(nameof(fromRow));
            if (fromCol <= 0) throw new ArgumentOutOfRangeException(nameof(fromCol));
            if (toRow <= 0) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toCol <= 0) throw new ArgumentOutOfRangeException(nameof(toCol));
            if (toRow < fromRow) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toCol < fromCol) throw new ArgumentOutOfRangeException(nameof(toCol));

            int rowCount = toRow - fromRow + 1;
            int colCount = toCol - fromCol + 1;

            if (values != null && (values.GetLength(0) != rowCount || values.GetLength(1) != colCount)) {
                throw new ArgumentException("Values array dimensions must match the specified range.", nameof(values));
            }

            var cells = new List<(int Row, int Column, object Value)>();
            for (int r = 0; r < rowCount; r++) {
                for (int c = 0; c < colCount; c++) {
                    object cellValue = values != null ? values[r, c] : string.Empty;
                    cells.Add((fromRow + r, fromCol + c, cellValue));
                }
            }

            Sheet.SetCellValuesParallel(cells);
            return this;
        }

        public SheetBuilder Column(Action<ColumnBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new ColumnBuilder(Sheet);
            action(builder);
            return this;
        }

        public SheetBuilder Table(Action<TableBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new TableBuilder(Sheet);
            action(builder);
            return this;
        }

        public SheetBuilder Style(Action<StyleBuilder> action) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            var builder = new StyleBuilder(Sheet);
            action(builder);
            return this;
        }

        public SheetBuilder AutoFilter(string range, Dictionary<uint, IEnumerable<string>>? criteria = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.AddAutoFilter(range, criteria);
            return this;
        }

        public SheetBuilder ConditionalColorScale(string range, SixLaborsColor startColor, SixLaborsColor endColor) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.AddConditionalColorScale(range, startColor, endColor);
            return this;
        }

        public SheetBuilder ConditionalDataBar(string range, SixLaborsColor color) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            Sheet.AddConditionalDataBar(range, color);
            return this;
        }

        public SheetBuilder AutoFit(bool columns, bool rows) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (columns) {
                Sheet.AutoFitColumns();
            }
            if (rows) {
                Sheet.AutoFitRows();
            }
            return this;
        }
    }
}
