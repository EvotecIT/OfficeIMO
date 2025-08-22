using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a row within a <see cref="WordTable"/> and exposes row level operations.
    /// </summary>
    public class WordTableRow {
        internal readonly TableRow _tableRow;

        /// <summary>
        /// Return all cells for given row
        /// </summary>
        public List<WordTableCell> Cells {
            get {
                var list = new List<WordTableCell>();
                foreach (TableCell cell in _tableRow.ChildElements.OfType<TableCell>().ToList()) {
                    WordTableCell wordCell = new WordTableCell(_document, _wordTable, this, cell);
                    list.Add(wordCell);
                }

                return list;
            }
        }
        /// <summary>
        /// Return first cell for given row
        /// </summary>
        public WordTableCell FirstCell => Cells.First();

        /// <summary>
        /// Return last cell for given row
        /// </summary>
        public WordTableCell LastCell => Cells.Last();

        /// <summary>
        /// Gets cells count
        /// </summary>
        public int CellsCount => Cells.Count;

        /// <summary>
        /// Gets or sets height of a row
        /// </summary>
        public int? Height {
            get {
                if (_tableRow.TableRowProperties != null) {
                    var rowHeight = _tableRow.TableRowProperties.OfType<TableRowHeight>().FirstOrDefault();
                    if (rowHeight?.Val != null) {
                        return (int)rowHeight.Val.Value;
                    }
                }
                return null;
            }
            set {
                if (value != null) {
                    AddTableRowProperties();
                    var tableRowProperties = _tableRow.TableRowProperties!;
                    var tableRowHeight = tableRowProperties.OfType<TableRowHeight>().FirstOrDefault();
                    if (tableRowHeight == null) {
                        tableRowHeight = new TableRowHeight();
                        tableRowProperties.InsertAt(tableRowHeight, 0);
                    }
                    tableRowHeight.Val = (uint)value;
                    tableRowHeight.HeightType = HeightRuleValues.Exact;
                } else {
                    var tableRowHeight = _tableRow.TableRowProperties?.OfType<TableRowHeight>().FirstOrDefault();
                    if (tableRowHeight != null) {
                        tableRowHeight.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the row is allowed to break across pages.
        /// When set to <c>false</c>, the row is kept intact on a single page.
        /// </summary>
        public bool AllowRowToBreakAcrossPages {
            get {
                if (_tableRow.TableRowProperties != null) {
                    var cantSplit = _tableRow.TableRowProperties.OfType<CantSplit>().FirstOrDefault();
                    if (cantSplit != null) {
                        return false;
                    }
                }

                return true;
            }
            set {
                if (value) {
                    if (_tableRow.TableRowProperties != null) {
                        var cantSplit = _tableRow.TableRowProperties.OfType<CantSplit>().FirstOrDefault();
                        if (cantSplit != null) {
                            cantSplit.Remove();
                        }
                    } else {
                        // nothing to do as CantSplit doesn't exists, because TableRowProperties doesn't exists
                        return;
                    }
                } else {
                    AddTableRowProperties();
                    var tableRowProperties = _tableRow.TableRowProperties!;
                    var cantSplit = tableRowProperties.OfType<CantSplit>().FirstOrDefault();
                    if (cantSplit == null) {
                        tableRowProperties.InsertAt(new CantSplit(), 0);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets header row at the top of each page
        /// Since this is a table row property, it is not possible to set it for a single row
        /// </summary>
        internal bool RepeatHeaderRowAtTheTopOfEachPage {
            get {
                if (_tableRow.TableRowProperties != null) {
                    var rowHeader = _tableRow.TableRowProperties.OfType<TableHeader>().FirstOrDefault();
                    if (rowHeader != null) {
                        return true;
                    }
                }
                return false;
            }
            set {
                AddTableRowProperties();
                var tableRowProperties = _tableRow.TableRowProperties!;
                var rowHeader = tableRowProperties.OfType<TableHeader>().FirstOrDefault();
                if (rowHeader != null) {
                    if (value == false) {
                        rowHeader.Remove();
                    }
                } else {
                    // Add table header
                    tableRowProperties.InsertAt(new TableHeader(), 0);

                }
            }
        }

        private readonly WordTable _wordTable;
        private readonly WordDocument _document;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableRow"/> class and creates an empty row.
        /// </summary>
        /// <param name="document">The parent <see cref="WordDocument"/>.</param>
        /// <param name="wordTable">The table to which this row belongs.</param>
        public WordTableRow(WordDocument document, WordTable wordTable) {
            // Create a row.
            TableRow tableRow = new TableRow();
            _tableRow = tableRow;
            _document = document;
            _wordTable = wordTable;

        }
        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableRow"/> class from an existing <see cref="TableRow"/>.
        /// </summary>
        /// <param name="wordTable">The parent <see cref="WordTable"/>.</param>
        /// <param name="row">The underlying Open XML table row.</param>
        /// <param name="document">The parent <see cref="WordDocument"/>.</param>
        public WordTableRow(WordTable wordTable, TableRow row, WordDocument document) {
            _document = document;
            _tableRow = row;
            _wordTable = wordTable;

            foreach (TableCell cell in row.ChildElements.OfType<TableCell>()) {
                WordTableCell wordCell = new WordTableCell(document, wordTable, this, cell);
            }
        }

        /// <summary>
        /// Appends the specified <see cref="WordTableCell"/> to the end of this row.
        /// </summary>
        /// <param name="cell">The cell to append.</param>
        public void Add(WordTableCell cell) {
            _tableRow.Append(cell._tableCell);
        }

        /// <summary>
        /// Remove a row
        /// </summary>
        public void Remove() {
            _tableRow.Remove();
        }

        /// <summary>
        /// Generate table row properties for the row if it doesn't exists
        /// </summary>
        internal void AddTableRowProperties() {
            if (_tableRow.TableRowProperties == null) {
                _tableRow.InsertAt(new TableRowProperties(), 0);
            }
        }

        /// <summary>
        /// Merges cells starting from the specified column across subsequent rows.
        /// </summary>
        /// <param name="cellIndex">Column index of the first cell to merge.</param>
        /// <param name="rowsCount">Number of rows below this one to merge.</param>
        /// <param name="copyParagraphs">True to move paragraphs from merged cells to the first cell; false to discard them.</param>
        public void MergeVertically(int cellIndex, int rowsCount, bool copyParagraphs = false) {
            var rows = _wordTable.Rows;
            int startIndex = rows.FindIndex(r => r._tableRow == _tableRow);
            if (startIndex < 0 || cellIndex >= CellsCount) {
                return;
            }

            var firstCell = rows[startIndex].Cells[cellIndex];
            firstCell.AddTableCellProperties();
            firstCell.VerticalMerge = MergedCellValues.Restart;
            var targetCell = firstCell._tableCell;

            for (int i = 0; i < rowsCount; i++) {
                int idx = startIndex + i + 1;
                if (idx >= rows.Count) break;

                var row = rows[idx];
                var cell = row.Cells[cellIndex];
                cell.AddTableCellProperties();

                if (copyParagraphs) {
                    var paragraphs = cell._tableCell.ChildElements.OfType<Paragraph>().ToList();
                    foreach (var paragraph in paragraphs) {
                        paragraph.Remove();
                        targetCell.Append(paragraph);
                    }
                    cell._tableCell.Append(new Paragraph());
                } else {
                    var paragraphs = cell._tableCell.ChildElements.OfType<Paragraph>().ToList();
                    foreach (var paragraph in paragraphs) {
                        paragraph.Remove();
                    }
                    cell._tableCell.Append(new Paragraph());
                }

                cell.VerticalMerge = MergedCellValues.Continue;
            }
        }
    }
}
