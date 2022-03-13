using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTable {
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
        /// <summary>
        /// Specifies that the first row conditional formatting shall be applied to the table.
        /// </summary>
        public bool? FirstRow {
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
        public bool? LastRow {
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
        public bool? FirstColumn {
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
        public bool? LastColumn {
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
        public bool? NoHorizontalBand {
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
        public bool? NoVerticalBand {
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

        private Table _table;

        private TableProperties _tableProperties {
            get {
                return _table.ChildElements.OfType<TableProperties>().FirstOrDefault();
            }
        }

        private WordDocument _document;
        //internal string Text;
        private WordSection _section;



        private Table GenerateTable(WordTableStyle tableStyle) {
            // Create an empty table.
            Table table = new Table();

            // Create a TableProperties object and specify its border information.
            //TableProperties tableProperties1 = new TableProperties(
            //    new TableBorders(
            //        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 }
            //    )
            //);

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = WordTableStyles.GetStyle(tableStyle);  //new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = tableStyle.ToString() };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);


            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tableProperties1);

            return table;
        }

        /// <summary>
        /// Used during load of the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="section"></param>
        /// <param name="table"></param>
        public WordTable(WordDocument document, WordSection section, Table table) {
            _table = table;
            _document = document;
            _section = section;

            foreach (TableRow row in table.ChildElements.OfType<TableRow>().ToList()) {
                WordTableRow tableRow = new WordTableRow(this, row, document);
            }
        }

        public WordTable(WordDocument document, WordSection section, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _section = section;
            _table = this.GenerateTable(tableStyle);

            for (int i = 0; i < rows; i++) {
                WordTableRow row = new WordTableRow(document, this);
                this._table.Append(row._tableRow);
                for (int j = 0; j < columns; j++) {
                    WordTableCell cell = new WordTableCell(document, this, row);
                }
            }

            // Append the table to the document.
            document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(_table);
        }

        public void AddRow(int cellsCount = 0) {
            WordTableRow row = new WordTableRow(_document, this);
            _table.Append(row._tableRow);
            AddCells(row, cellsCount);
        }

        private void AddCells(WordTableRow row, int cellsCount = 0) {
            if (cellsCount == 0) {
                // we try to get the last row and fill it with same number of cells
                cellsCount = this.Rows[this.RowsCount - 2].CellsCount;
            }
            for (int j = 0; j < cellsCount; j++) {
                WordTableCell cell = new WordTableCell(_document, this, row);
            }
        }

        public void AddRow(int rowsCount, int cellsCount) {
            for (int i = 0; i < rowsCount; i++) {
                AddRow(cellsCount);
            }
        }
        private void Add(WordTableRow row) {
            _table.Append(row._tableRow);
            // this.Rows.Add(row);
        }

        public void Remove() {
            _table.Remove();
        }
    }
}
