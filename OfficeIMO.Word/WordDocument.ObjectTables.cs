using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;

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

            var columns = GetColumnNames(first);
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
                    var value = GetValue(rowObj, columns[c]);
                    var paragraph = GetOrCreateParagraph(table, rowIndex, c);
                    paragraph.Text = value?.ToString() ?? string.Empty;
                }

                rowIndex++;
            }

            return table;
        }

        private static IReadOnlyList<string> GetColumnNames(object item) {
            if (item is IReadOnlyDictionary<string, object?> roDict) {
                return roDict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            }

            if (item is IDictionary<string, object?> dict) {
                return dict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            }

            if (item is IDictionary legacyDict) {
                var names = new List<string>();
                foreach (DictionaryEntry entry in legacyDict) {
                    var key = entry.Key?.ToString();
                    if (!string.IsNullOrWhiteSpace(key)) {
                        names.Add(key!);
                    }
                }
                return names;
            }

            var props = item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
                .OrderBy(p => p.MetadataToken)
                .Select(p => p.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            return props;
        }

        private static object? GetValue(object item, string column) {
            if (item is IReadOnlyDictionary<string, object?> roDict) {
                return roDict.TryGetValue(column, out var value) ? value : null;
            }

            if (item is IDictionary<string, object?> dict) {
                return dict.TryGetValue(column, out var value) ? value : null;
            }

            if (item is IDictionary legacyDict) {
                if (legacyDict.Contains(column)) {
                    return legacyDict[column];
                }

                foreach (DictionaryEntry entry in legacyDict) {
                    var key = entry.Key?.ToString();
                    if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase)) {
                        return entry.Value;
                    }
                }

                return null;
            }

            var prop = item.GetType().GetProperty(column, BindingFlags.Public | BindingFlags.Instance);
            return prop?.GetValue(item);
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
