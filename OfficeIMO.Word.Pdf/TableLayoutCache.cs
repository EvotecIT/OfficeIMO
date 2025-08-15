using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    internal static class TableLayoutCache {
        private static readonly ConditionalWeakTable<WordTable, TableLayout> _cache = new();

        internal static TableLayout GetLayout(WordTable table) {
            if (_cache.TryGetValue(table, out TableLayout layout)) {
                return layout;
            }

            List<IReadOnlyList<WordTableCell>> rows = TableBuilder.Map(table).ToList();
            int columnCount = rows.Max(r => r.Count);
            float[] widths = new float[columnCount];

            foreach (IReadOnlyList<WordTableCell> row in rows) {
                for (int i = 0; i < row.Count; i++) {
                    WordTableCell cell = row[i];
                    float width = 0f;
                    if (cell.Width.HasValue && cell.WidthType == TableWidthUnitValues.Dxa) {
                        width = cell.Width.Value / 20f;
                    }

                    if (cell.HasNestedTables) {
                        foreach (WordTable nested in cell.NestedTables) {
                            TableLayout nestedLayout = GetLayout(nested);
                            float nestedWidth = nestedLayout.ColumnWidths.Sum();
                            if (nestedWidth > width) {
                                width = nestedWidth;
                            }
                        }
                    }

                    if (width > widths[i]) {
                        widths[i] = width;
                    }
                }
            }

            layout = new TableLayout(rows, widths);
            _cache.Add(table, layout);
            return layout;
        }
    }
}

