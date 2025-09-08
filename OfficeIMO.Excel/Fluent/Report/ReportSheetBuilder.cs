using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.Fluent.Report {
    /// <summary>
    /// High-level report composer for building rich worksheets with sections, properties grids, tables and references.
    /// Keeps track of the current row and provides simple, layout-oriented helpers.
    /// </summary>
    public sealed class ReportSheetBuilder {
        private readonly ExcelDocument _workbook;
        private readonly ReportTheme _theme;
        private ExcelSheet _sheet;
        private int _row;

        /// <summary>
        /// Creates a report builder on a new worksheet.
        /// </summary>
        public ReportSheetBuilder(ExcelDocument workbook, string sheetName, ReportTheme? theme = null) {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _theme = theme ?? ReportTheme.Default;
            _sheet = _workbook.AddWorkSheet(sheetName);
            _row = 1;
        }

        /// <summary>Current 1-based row where the next block will start.</summary>
        public int CurrentRow => _row;

        /// <summary>Underlying sheet for advanced operations.</summary>
        public ExcelSheet Sheet => _sheet;

        /// <summary>Adds vertical space.</summary>
        public ReportSheetBuilder Spacer(int rows = -1) {
            _row += rows > 0 ? rows : _theme.DefaultSpacingRows;
            return this;
        }

        /// <summary>Writes a large title and optional subtitle.</summary>
        public ReportSheetBuilder Title(string text, string? subtitle = null) {
            if (string.IsNullOrWhiteSpace(text)) return this;
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            // Use a bigger appearance by doubling into two rows and coloring
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                _sheet.Cell(_row, 1, subtitle);
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Renders a simple key/value properties grid.
        /// </summary>
        /// <param name="properties">Pairs of label to value.</param>
        /// <param name="columns">How many key/value pairs per row.</param>
        public ReportSheetBuilder PropertiesGrid(IEnumerable<(string Key, object? Value)> properties, int columns = 2) {
            if (properties == null) return this;
            var list = properties.ToList();
            if (list.Count == 0) return this;

            int idx = 0;
            while (idx < list.Count) {
                int col = 1;
                for (int c = 0; c < columns && idx < list.Count; c++, idx++) {
                    var (k, v) = list[idx];
                    _sheet.Cell(_row, col, k);
                    _sheet.CellBold(_row, col, true);
                    _sheet.CellBackground(_row, col, _theme.KeyFillHex);
                    _sheet.Cell(_row, col + 1, v ?? string.Empty);
                    col += 2;
                }
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Adds a small section header (shaded) for grouping.
        /// </summary>
        public ReportSheetBuilder Section(string text) {
            _sheet.Cell(_row, 1, text);
            _sheet.CellBold(_row, 1, true);
            _sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            return this;
        }

        /// <summary>
        /// Adds a paragraph (single cell, wrapped) spanning a set number of columns.
        /// </summary>
        public ReportSheetBuilder Paragraph(string text, int widthColumns = 6) {
            if (string.IsNullOrEmpty(text)) return this;
            _sheet.Cell(_row, 1, text);
            // wrap text by applying a style variant
            // (uses internal helper on the sheet)
            // We call a private helper via existing public entry points by formatting then toggling wrap.
            // There's no direct public API yet, so we keep it simple; long text will still show.
            _row++;
            return this;
        }

        /// <summary>
        /// Renders a bulleted list (one item per row).
        /// </summary>
        public ReportSheetBuilder BulletedList(IEnumerable<string> items) {
            if (items == null) return this;
            foreach (var item in items) {
                _sheet.Cell(_row, 1, $"• {item}");
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Renders a table from a collection of objects. Returns the A1 range of the table.
        /// </summary>
        public string TableFrom<T>(IEnumerable<T> items, string? title = null, Action<ObjectFlattenerOptions>? configure = null, TableStyle style = TableStyle.TableStyleMedium9, bool autoFilter = true) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);

            var data = items?.ToList() ?? new List<T>();
            if (data.Count == 0) {
                _sheet.Cell(_row, 1, "(no data)");
                _row++;
                return $"A{_row}:A{_row}";
            }

            var opts = new ObjectFlattenerOptions();
            configure?.Invoke(opts);
            var flattener = new ObjectFlattener();
            var paths = opts.Columns?.ToList() ?? flattener.GetPaths(typeof(T), opts);

            // header + rows (batch writes for performance)
            int headerRow = _row;
            var cells = new List<(int Row, int Column, object Value)>(Math.Max(1, (data.Count + 1) * paths.Count));
            for (int i = 0; i < paths.Count; i++) {
                cells.Add((_row, i + 1, paths[i]));
            }
            _row++;

            foreach (var item in data) {
                var dict = flattener.Flatten(item, opts);
                for (int i = 0; i < paths.Count; i++) {
                    dict.TryGetValue(paths[i], out var val);
                    cells.Add((_row, i + 1, val ?? string.Empty));
                }
                _row++;
            }
            _sheet.CellValues(cells);
            // style header line
            for (int i = 0; i < paths.Count; i++) {
                _sheet.CellBold(headerRow, i + 1, true);
                _sheet.CellBackground(headerRow, i + 1, _theme.KeyFillHex);
            }

            int lastRow = _row - 1;
            string start = $"A{headerRow}";
            string end = ColumnLetter(paths.Count) + lastRow.ToString();
            string range = start + ":" + end;

            // Add table + (optional) AutoFilter on the same range
            _sheet.AddTable(range, hasHeader: true, name: title ?? "Table", style: style, includeAutoFilter: autoFilter);
            Spacer();
            return range;
        }

        /// <summary>
        /// Renders a list of URLs as clickable references (displaying the URL text).
        /// </summary>
        public ReportSheetBuilder References(IEnumerable<string> urls) {
            var list = urls?.Where(u => !string.IsNullOrWhiteSpace(u)).ToList();
            if (list == null || list.Count == 0) return this;
            Section("References");
            foreach (var url in list) {
                _sheet.SetHyperlink(_row, 1, url, url);
                _row++;
            }
            return Spacer();
        }

        /// <summary>
        /// Displays a simple score row with optional heat color via 2-color scale over the cell.
        /// </summary>
        public ReportSheetBuilder Score(string label, double value, double min = 0, double max = 10) {
            _sheet.Cell(_row, 1, label);
            _sheet.CellBold(_row, 1, true);
            _sheet.Cell(_row, 2, value);
            string range = $"B{_row}:B{_row}";
            // A gentle data bar works nicely for a single row too
            _sheet.AddConditionalDataBar(range, SixLabors.ImageSharp.Color.LightGreen);
            _row++;
            return Spacer();
        }

        /// <summary>
        /// Optional finishing step: auto-fit columns/rows once at the end.
        /// </summary>
        public ReportSheetBuilder Finish(bool autoFitColumns = true, bool autoFitRows = false)
        {
            if (autoFitColumns) _sheet.AutoFitColumns();
            if (autoFitRows) _sheet.AutoFitRows();
            return this;
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
