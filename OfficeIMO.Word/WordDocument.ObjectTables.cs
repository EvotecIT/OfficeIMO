using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Adds a table based on a sequence of objects by projecting properties or dictionary keys into columns.
        /// </summary>
        /// <param name="items">Objects to insert into the table.</param>
        /// <param name="style">Table style to apply.</param>
        /// <param name="includeHeader">Whether to include a header row.</param>
        /// <param name="layout">Optional table layout (Autofit or Fixed).</param>
        /// <returns>The created <see cref="WordTable"/>.</returns>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
        public WordTable AddTableFromObjects(IEnumerable<object?> items, WordTableStyle style = WordTableStyle.TableGrid, bool includeHeader = true, TableLayoutValues? layout = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var list = items.ToList();
            if (list.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            var first = list.FirstOrDefault();
            if (first == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            var columns = ObjectDataHelpers.GetColumnNames(first);
            if (columns.Count == 0) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            int rows = list.Count + (includeHeader ? 1 : 0);
            int cols = columns.Count;

            var table = AddTable(rows, cols, style);

            if (layout.HasValue) {
                table.LayoutType = layout.Value;
            }

            int rowIndex = 0;
            if (includeHeader) {
                for (int c = 0; c < cols; c++) {
                    var headerParagraph = GetOrCreateParagraph(table, 0, c);
                    headerParagraph.Text = columns[c];
                }
                rowIndex = 1;
            }

            for (int r = 0; r < list.Count; r++) {
                var rowObj = list[r];
                if (rowObj == null) {
                    throw new InvalidOperationException("Data rows cannot contain null entries.");
                }

                for (int c = 0; c < cols; c++) {
                    var value = ObjectDataHelpers.GetValue(rowObj, columns[c]);
                    var paragraph = GetOrCreateParagraph(table, rowIndex, c);
                    paragraph.Text = value?.ToString() ?? string.Empty;
                }

                rowIndex++;
            }

            return table;
        }

        private static WordParagraph GetOrCreateParagraph(WordTable table, int rowIndex, int columnIndex) {
            var rows = table.Rows ?? throw new InvalidOperationException("Table rows collection is missing.");
            var row = rows[rowIndex];
            var cells = row.Cells ?? throw new InvalidOperationException("Table cells collection is missing.");
            var cell = cells[columnIndex];
            return cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
        }
    }
}
