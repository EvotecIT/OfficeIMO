using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Simple multi-column layout for placing lightweight blocks side by side.
    /// </summary>
    public sealed partial class SheetComposer
    {
        public sealed class ColumnComposer
        {
            private readonly ExcelSheet _sheet;
            private readonly SheetTheme _theme;
            private readonly int _baseCol;
            private readonly int _startRow;
            private int _row;

            internal ColumnComposer(ExcelSheet sheet, SheetTheme theme, int startRow, int baseCol)
            { _sheet = sheet; _theme = theme; _startRow = startRow; _row = startRow; _baseCol = baseCol; }

            public int RowsUsed => _row - _startRow;

            public ColumnComposer Spacer(int rows = 1) { _row += Math.Max(1, rows); return this; }

            public ColumnComposer Section(string text)
            {
                _sheet.Cell(_row, _baseCol, text);
                _sheet.CellBold(_row, _baseCol, true);
                _sheet.CellBackground(_row, _baseCol, _theme.SectionHeaderFillHex);
                _row++;
                return this;
            }

            public ColumnComposer Paragraph(string text)
            {
                if (!string.IsNullOrEmpty(text)) { _sheet.Cell(_row, _baseCol, text); _row++; }
                return this;
            }

            public ColumnComposer BulletedList(IEnumerable<string> items)
            {
                if (items == null) return this;
                foreach (var item in items) { _sheet.Cell(_row, _baseCol, $"• {item}"); _row++; }
                return this;
            }

            public ColumnComposer KeyValue(string key, object? value)
            {
                _sheet.Cell(_row, _baseCol, key);
                _sheet.CellBold(_row, _baseCol, true);
                _sheet.CellBackground(_row, _baseCol, _theme.KeyFillHex);
                _sheet.Cell(_row, _baseCol + 1, value ?? string.Empty);
                _row++;
                return this;
            }

            public ColumnComposer KeyValues(IEnumerable<(string Key, object? Value)> pairs)
            {
                if (pairs == null) return this;
                foreach (var (k, v) in pairs) KeyValue(k, v);
                return this;
            }
        }

        /// <summary>
        /// Places N columns side-by-side starting at the current row. Each action receives a ColumnComposer
        /// scoped to its own column. The main composer advances to the maximum height used by the columns.
        /// </summary>
        /// <param name="count">Number of columns (2–4 recommended).</param>
        /// <param name="configure">Callback that receives an array of ColumnComposer objects.</param>
        /// <param name="columnWidth">Width per column in grid columns (for relative positioning only).</param>
        /// <param name="gutter">Spacing between columns in grid columns.</param>
        public SheetComposer Columns(int count, Action<ColumnComposer[]> configure, int columnWidth = 3, int gutter = 1)
        {
            if (count <= 1) return this;
            int startRow = _row;
            var cols = new ColumnComposer[count];
            int baseCol = 1;
            for (int i = 0; i < count; i++)
            {
                cols[i] = new ColumnComposer(Sheet, _theme, startRow, baseCol);
                baseCol += columnWidth + gutter;
            }
            configure?.Invoke(cols);
            int maxRows = 0; foreach (var c in cols) if (c.RowsUsed > maxRows) maxRows = c.RowsUsed;
            _row = startRow + maxRows;
            return Spacer();
        }
    }
}
