using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Adds a table from runtime objects by projecting their public properties or dictionary keys into columns.
        /// </summary>
        /// <param name="items">Objects to insert into the table. Column names come from the first item.</param>
        /// <param name="options">Common column selection, ordering, header, formatting, alignment, name and caption options. Pass <see cref="WordTableOptions"/> to choose the table style and layout mode.</param>
        /// <returns>The created <see cref="WordTable"/>.</returns>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, use AddTable<T> with declared rows or pre-flatten items.")]
#pragma warning disable RS0026 // Same shape as MarkdownDocument.AddTable(object?, OfficeTableOptions?); a required options parameter would instead break RS0027 for the shipped typed overloads.
        public WordTable AddTable(IEnumerable<object?> items, OfficeTableOptions? options = null) {
#pragma warning restore RS0026
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            ObjectDataHelpers.ValidateTableAlignments(options);
            var list = items.ToList();
            if (list.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }
            if (list.Any(item => item == null)) {
                throw new ArgumentException("Data rows cannot contain null entries.", nameof(items));
            }

            var columns = ObjectDataHelpers.SelectTableColumns(ObjectDataHelpers.GetColumnNames(list[0]!), options);
            if (columns.Length == 0) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries, and check the Include and Exclude options.");
            }

            bool includeHeader = options?.IncludeHeaders ?? true;
            var table = CreateOptionsTable(list.Count + (includeHeader ? 1 : 0), columns.Length, options);

            int rowIndex = 0;
            if (includeHeader) {
                for (int c = 0; c < columns.Length; c++) {
                    GetOrCreateParagraph(table, 0, c).Text = ObjectDataHelpers.GetTableHeader(columns[c], options);
                }
                rowIndex = 1;
            }

            foreach (var rowObj in list) {
                for (int c = 0; c < columns.Length; c++) {
                    var value = ObjectDataHelpers.FormatTableValue(columns[c], ObjectDataHelpers.GetValue(rowObj!, columns[c]), options);
                    GetOrCreateParagraph(table, rowIndex, c).Text = value?.ToString() ?? string.Empty;
                }

                rowIndex++;
            }

            ApplyTableAlignments(table, options);
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
