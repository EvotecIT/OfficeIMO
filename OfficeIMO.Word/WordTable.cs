using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable : WordElement {
        public List<WordParagraph> Paragraphs {
            get {
                List<WordParagraph> list = new List<WordParagraph>();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        list.AddRange(cell.Paragraphs);
                    }
                }
                return list;
            }
        }

        public WordTableStyle? Style {
            get {
                if (_tableProperties != null && _tableProperties.TableStyle != null) {
                    var style = _tableProperties.TableStyle.Val;
                    return WordTableStyles.GetStyle(style);
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableStyle != null && value != null) {
                    _tableProperties.TableStyle = WordTableStyles.GetStyle(value.Value);
                }
            }
        }

        public TableRowAlignmentValues? Alignment {
            get {
                if (_tableProperties != null && _tableProperties.TableJustification != null) {
                    return _tableProperties.TableJustification.Val;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableJustification == null) {
                    _tableProperties.TableJustification = new TableJustification();
                }
                if (value != null) {
                    _tableProperties.TableJustification.Val = value.Value;
                } else {
                    _tableProperties.TableJustification.Remove();
                }
            }
        }

        public TableWidthUnitValues? WidthType {
            get {
                if (_tableProperties != null && _tableProperties.TableWidth != null) {
                    return _tableProperties.TableWidth.Type;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableWidth == null) {
                    if (value == TableWidthUnitValues.Auto) {
                        _tableProperties.TableWidth = new TableWidth() {
                            Type = value,
                            Width = "0"
                        };
                    } else {
                        _tableProperties.TableWidth = new TableWidth() {
                            Type = value,
                            Width = "5000"
                        };
                    }
                } else {
                    if (value == TableWidthUnitValues.Auto) {
                        _tableProperties.TableWidth.Type = value;
                        _tableProperties.TableWidth.Width = "0";
                    } else {
                        _tableProperties.TableWidth.Type = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets width of a table
        /// </summary>
        public int? Width {
            get {
                if (_tableProperties != null && _tableProperties.TableWidth != null) {
                    if (_tableProperties.TableWidth.Width != null) {
                        return int.Parse(_tableProperties.TableWidth.Width);
                    }
                }
                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableWidth == null) {
                    _tableProperties.TableWidth = new TableWidth() {
                        Type = TableWidthUnitValues.Pct,
                        Width = value.ToString()
                    };
                } else {
                    _tableProperties.TableWidth.Width = value.ToString();
                }
            }
        }

        /// <summary>
        /// Gets or sets layout of a table
        /// </summary>
        public TableLayoutValues? LayoutType {
            get {
                if (_tableProperties != null && _tableProperties.TableLayout != null) {
                    return _tableProperties.TableLayout.Type;
                }
                return TableLayoutValues.Autofit;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableLayout == null) {
                    _tableProperties.TableLayout = new TableLayout();
                }
                if (value != null) {
                    _tableProperties.TableLayout.Type = value;
                } else {
                    _tableProperties.TableLayout.Remove();
                }
            }
        }

        /// <summary>
        /// Specifies that the first row conditional formatting shall be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingFirstRow {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.FirstRow;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.FirstRow = value;
                }
            }
        }
        /// <summary>
        /// Specifies that the last row conditional formatting shall be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingLastRow {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.LastRow;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.LastRow = value;
                }
            }
        }
        /// <summary>
        /// Specifies that the first column conditional formatting shall be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingFirstColumn {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.FirstColumn;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.FirstColumn = value;
                }
            }
        }
        /// <summary>
        /// Specifies that the last column conditional formatting shall be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingLastColumn {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.LastColumn;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.LastColumn = value;
                }
            }
        }
        /// <summary>
        /// Specifies that the horizontal banding conditional formatting shall not be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingNoHorizontalBand {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.NoHorizontalBand;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.NoHorizontalBand = value;
                }
            }
        }
        /// <summary>
        /// Specifies that the vertical banding conditional formatting shall not be applied to the table.
        /// </summary>
        public bool? ConditionalFormattingNoVerticalBand {
            get {
                if (_tableProperties != null && _tableProperties.TableLook != null) {
                    return _tableProperties.TableLook.NoVerticalBand;
                }
                return null;
            }
            set {
                if (_tableProperties != null && _tableProperties.TableLook != null && value != null) {
                    _tableProperties.TableLook.NoVerticalBand = value;
                }
            }
        }

        /// <summary>
        /// Specifies that the first row shall be repeated at the top of each page on which the table is displayed.
        /// </summary>
        public bool RepeatHeaderRowAtTheTopOfEachPage {
            get => Rows[0].RepeatHeaderRowAtTheTopOfEachPage;
            set => Rows[0].RepeatHeaderRowAtTheTopOfEachPage = value;
        }

        public int RowsCount => this.Rows.Count;

        public List<WordTableRow> Rows {
            get {
                var list = new List<WordTableRow>();

                foreach (TableRow row in _table.ChildElements.OfType<TableRow>()) {
                    WordTableRow tableRow = new WordTableRow(this, row, _document);
                    list.Add(tableRow);
                }

                return list;
            }
        }

        public WordTableRow FirstRow {
            get {
                return Rows.First();
            }
        }
        public WordTableRow LastRow {
            get {
                return Rows.Last();
            }
        }

        internal Table _table;

        internal TableProperties _tableProperties {
            get {
                return _table.ChildElements.OfType<TableProperties>().FirstOrDefault();
            }
        }

        private WordDocument _document;
        //internal string Text;
        //private WordSection _section;

        private Header _header {
            get {
                var parent = _table.Parent;
                if (parent is Header) {
                    return (Header)parent;
                }

                return null;
            }
        }

        private Footer _footer {
            get {
                var parent = _table.Parent;
                if (parent is Footer) {
                    return (Footer)parent;
                }

                return null;
            }
        }

        public WordTablePosition Position;


        private Table GenerateTable(WordDocument document, int rows, int columns, WordTableStyle tableStyle) {
            Table table = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = WordTableStyles.GetStyle(tableStyle);
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);

            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tableProperties1);

            TableGrid tableGrid1 = new TableGrid();
            for (int i = 0; i < columns; i++) {
                GridColumn gridColumn1 = new GridColumn() { };
                tableGrid1.Append(gridColumn1);
            }
            table.Append(tableGrid1);

            for (int i = 0; i < rows; i++) {
                WordTableRow row = new WordTableRow(document, this);
                table.Append(row._tableRow);
                for (int j = 0; j < columns; j++) {
                    WordTableCell cell = new WordTableCell(document, this, row);
                }
            }
            return table;
        }

        /// <summary>
        /// Used during load of the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="table"></param>
        internal WordTable(WordDocument document, Table table) {
            _table = table;
            _document = document;

            foreach (TableRow row in table.ChildElements.OfType<TableRow>().ToList()) {
                WordTableRow tableRow = new WordTableRow(this, row, document);
            }

            // Establish Position property
            Position = new WordTablePosition(this);
        }

        internal WordTable(WordDocument document, WordParagraph wordParagraph, int rows, int columns, WordTableStyle tableStyle, string location) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            if (location == "After") {
                // Append the table to the document after given paragraph
                wordParagraph._paragraph.InsertAfterSelf(_table);
            } else {
                // Append the table to the document before given paragraph
                wordParagraph._paragraph.InsertBeforeSelf(_table);
            }
        }

        internal WordTable(WordDocument document, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            // Append the table to the document.
            document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(_table);
        }

        public WordTable(WordDocument document, TableCell wordTableCell, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;

            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            wordTableCell.Append(_table);
        }

        internal WordTable(WordDocument document, Footer footer, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            footer.Append(_table);
        }
        internal WordTable(WordDocument document, Header header, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            header.Append(_table);
        }

        /// <summary>
        /// Add row to an existing table with the specified number of columns
        /// </summary>
        /// <param name="cellsCount"></param>
        public WordTableRow AddRow(int cellsCount = 0) {
            WordTableRow row = new WordTableRow(_document, this);
            _table.Append(row._tableRow);
            AddCells(row, cellsCount);
            return row;
        }

        /// <summary>
        /// Add cells to an existing row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellsCount"></param>
        private void AddCells(WordTableRow row, int cellsCount = 0) {
            if (cellsCount == 0) {
                // we try to get the last row and fill it with same number of cells
                cellsCount = this.Rows[this.RowsCount - 2].CellsCount;
            }
            for (int j = 0; j < cellsCount; j++) {
                WordTableCell cell = new WordTableCell(_document, this, row);
            }
        }

        /// <summary>
        /// Add specified number of rows to an existing table with the specified number of columns
        /// </summary>
        /// <param name="rowsCount"></param>
        /// <param name="cellsCount"></param>
        public List<WordTableRow> AddRow(int rowsCount, int cellsCount) {
            List<WordTableRow> rows = new List<WordTableRow>();
            for (int i = 0; i < rowsCount; i++) {
                rows.Add(AddRow(cellsCount));
            }
            return rows;
        }

        /// <summary>
        /// Adds a row to the table using WordTableRow object
        /// </summary>
        /// <param name="row"></param>
        private void AddRow(WordTableRow row) {
            _table.Append(row._tableRow);
        }

        /// <summary>
        /// Remove table from document
        /// </summary>
        public void Remove() {
            _table.Remove();
        }

        /// <summary>
        /// Generate table properties for the table if it doesn't exists
        /// </summary>
        internal void CheckTableProperties() {
            if (_tableProperties == null) {
                _table.AppendChild(new TableProperties());
            }
        }
    }
}
