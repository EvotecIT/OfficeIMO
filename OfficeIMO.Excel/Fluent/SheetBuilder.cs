using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using OfficeIMO.Excel.Utilities;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel.Fluent {
    public class SheetBuilder {
        private readonly ExcelFluentWorkbook _fluent;
        internal ExcelSheet? Sheet { get; private set; }
        private int _currentRow = 1;
        private string? _lastRange;

        internal SheetBuilder(ExcelFluentWorkbook fluent) {
            _fluent = fluent;
        }

        public SheetBuilder AddSheet(string name = "") {
            Sheet = _fluent.Workbook.AddWorkSheet(name);
            return this;
        }

        public SheetBuilder HeaderRow(params object?[] values) {
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

            var enumerable = data.ToList();
            if (!enumerable.Any()) return this;

            var paths = options.Columns?.ToList() ?? flattener.GetPaths(typeof(T), options);
            var headers = paths.Select(p => TransformHeader(p, options)).ToList();
            int startRow = _currentRow;
            HeaderRow(headers.Cast<object?>().ToArray());

            int dataRows = 0;
            foreach (var item in enumerable) {
                var dict = flattener.Flatten(item, options);
                if (options.CollectionMode == CollectionMode.ExpandRows) {
                    var collectionPath = paths.FirstOrDefault(p => dict.TryGetValue(p, out var val) && val is IEnumerable && val is not string);
                    if (collectionPath != null && dict[collectionPath] is IEnumerable coll) {
                        var list = coll.Cast<object?>().ToList();
                        if (list.Count == 0) {
                            Row(r => r.Values(paths.Select(p => dict.TryGetValue(p, out var v) ? v : null).ToArray()));
                            dataRows++;
                        } else {
                            foreach (var element in list) {
                                var rowValues = paths.Select(p => p == collectionPath ? element : dict.TryGetValue(p, out var v) ? v : (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray();
                                Row(r => r.Values(rowValues));
                                dataRows++;
                            }
                        }
                        continue;
                    }
                }

                Row(r => r.Values(paths.Select(p => dict.TryGetValue(p, out var v) ? v : (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray()));
                dataRows++;
            }

            int endRow = startRow + dataRows;
            _lastRange = $"A{startRow}:{ColumnLetter(headers.Count)}{endRow}";

            return this;
        }

        public SheetBuilder Table(string name, Action<TableBuilder>? configure = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (string.IsNullOrEmpty(_lastRange)) throw new InvalidOperationException("RowsFrom must be called before Table");
            var builder = new TableBuilder(Sheet);
            configure?.Invoke(builder);
            builder.Build(_lastRange!, name);
            return this;
        }

        public SheetBuilder Cell(int row, int column, object? value = null, string? formula = null, string? numberFormat = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
            Sheet.Cell(row, column, value, formula, numberFormat);
            return this;
        }

        public SheetBuilder Range(int fromRow, int fromCol, int toRow, int toCol, object[,]? values = null) {
            if (Sheet == null) throw new InvalidOperationException("Sheet not initialized");
            if (fromRow < 1) throw new ArgumentOutOfRangeException(nameof(fromRow));
            if (fromCol < 1) throw new ArgumentOutOfRangeException(nameof(fromCol));
            if (toRow < 1) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toCol < 1) throw new ArgumentOutOfRangeException(nameof(toCol));
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

            Sheet.CellValuesParallel(cells);
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

        private static string TransformHeader(string path, ObjectFlattenerOptions opts) {
            foreach (var prefix in opts.HeaderPrefixTrimPaths) {
                if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    path = path.Substring(prefix.Length);
                }
            }
            return opts.HeaderCase switch {
                HeaderCase.Pascal => string.Concat(path.Split('.').Select(s => char.ToUpperInvariant(s[0]) + s.Substring(1))),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

        private static string ColumnLetter(int column) {
            var dividend = column;
            var columnName = string.Empty;
            while (dividend > 0) {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
    }
}
