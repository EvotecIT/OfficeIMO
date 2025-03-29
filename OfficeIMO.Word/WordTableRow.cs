using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
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
                    if (rowHeight != null) {
                        return (int)rowHeight.Val.Value;
                    }
                }
                return null;
            }
            set {
                if (value != null) {
                    AddTableRowProperties();
                    var tableRowHeight = _tableRow.TableRowProperties.OfType<TableRowHeight>().FirstOrDefault();
                    if (tableRowHeight == null) {
                        _tableRow.TableRowProperties.InsertAt(new TableRowHeight(), 0);
                        tableRowHeight = _tableRow.TableRowProperties.OfType<TableRowHeight>().FirstOrDefault();
                    }
                    tableRowHeight.Val = (uint)value;
                } else {
                    var tableRowHeight = _tableRow.TableRowProperties.OfType<TableRowHeight>().FirstOrDefault();
                    if (tableRowHeight != null) {
                        tableRowHeight.Remove();
                    }
                }
            }
        }


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
                    var cantSplit = _tableRow.TableRowProperties.OfType<CantSplit>().FirstOrDefault();
                    if (cantSplit == null) {
                        _tableRow.TableRowProperties.InsertAt(new CantSplit(), 0);
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
                var rowHeader = _tableRow.TableRowProperties.OfType<TableHeader>().FirstOrDefault();
                if (rowHeader != null) {
                    if (value == false) {
                        rowHeader.Remove();
                    }
                } else {
                    // Add table header
                    _tableRow.TableRowProperties.InsertAt(new TableHeader(), 0);

                }
            }
        }

        private readonly WordTable _wordTable;
        private readonly WordDocument _document;

        public WordTableRow(WordDocument document, WordTable wordTable) {
            // Create a row.
            TableRow tableRow = new TableRow();
            _tableRow = tableRow;
            _document = document;
            _wordTable = wordTable;

        }
        public WordTableRow(WordTable wordTable, TableRow row, WordDocument document) {
            _document = document;
            _tableRow = row;
            _wordTable = wordTable;

            foreach (TableCell cell in row.ChildElements.OfType<TableCell>()) {
                WordTableCell wordCell = new WordTableCell(document, wordTable, this, cell);
            }
        }

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
    }
}
