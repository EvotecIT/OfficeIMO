using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a table within a <see cref="WordDocument"/>.
    /// </summary>
    public partial class WordTable : WordElement {
        /// <summary>
        /// Gets all <see cref="WordParagraph"/> instances contained in the table.
        /// </summary>
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

        /// <summary>
        /// Allow row to break across pages
        /// This sets each row to allow page break
        /// You can set each row separately as well
        /// Getting value returns true if any row allows page break
        /// For complete control use WordTableRow.AllowRowToBreakAcrossPages
        /// </summary>
        public bool AllowRowToBreakAcrossPages {
            get {
                bool allowRowToBreakAcrossPages = false;
                foreach (var row in this.Rows) {
                    if (row.AllowRowToBreakAcrossPages) {
                        allowRowToBreakAcrossPages = true;
                        break;
                    }
                }
                return allowRowToBreakAcrossPages;
            }
            set {
                foreach (var row in this.Rows) {
                    row.AllowRowToBreakAcrossPages = value;
                }
            }
        }

        /// <summary>
        /// Allow Header to repeat on each page
        /// This applies to only header of the table (first row)
        /// </summary>
        public bool RepeatAsHeaderRowAtTheTopOfEachPage {
            get {
                var tableRowProperties = this.Rows[0]._tableRow.TableRowProperties;
                var tableHeader = tableRowProperties?.OfType<TableHeader>().FirstOrDefault();
                return tableHeader != null;
            }
            set {
                if (value) {
                    this.Rows[0].AddTableRowProperties();
                    var tableProperties = this.Rows[0]._tableRow.TableRowProperties;
                    var tableHeader = tableProperties?.OfType<TableHeader>().FirstOrDefault();
                    if (tableHeader == null) {
                        tableProperties?.InsertAt(new TableHeader(), 0);
                    }
                } else {
                    var tableRowTableRowProperties = this.Rows[0]._tableRow.TableRowProperties;
                    var tableHeader = tableRowTableRowProperties?.OfType<TableHeader>().FirstOrDefault();
                    if (tableHeader != null) {
                        tableRowTableRowProperties!.RemoveChild(tableHeader);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the built-in table style applied to this table.
        /// Returns <c>null</c> when no style is assigned.
        /// </summary>
        public WordTableStyle? Style {
            get {
                if (_tableProperties != null && _tableProperties.TableStyle != null) {
                    var styleValue = _tableProperties.TableStyle.Val?.Value;
                    if (styleValue is { Length: > 0 } s) {
                        return WordTableStyles.GetStyle(s);
                    }
                }
                return null;
            }
            set {
                if (_tableProperties?.TableStyle != null && value != null) {
                    _tableProperties.TableStyle = WordTableStyles.GetStyle(value.Value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the horizontal alignment of the table within the page.
        /// </summary>
        public TableRowAlignmentValues? Alignment {
            get {
                if (_tableProperties != null && _tableProperties.TableJustification != null) {
                    return _tableProperties.TableJustification.Val?.Value;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableJustification == null) {
                    _tableProperties.TableJustification = new TableJustification();
                }
                if (value != null) {
                    _tableProperties.TableJustification.Val = value.Value;
                } else {
                    _tableProperties.TableJustification.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the unit type used by the <see cref="Width"/> property.
        /// </summary>
        public TableWidthUnitValues? WidthType {
            get {
                if (_tableProperties != null && _tableProperties.TableWidth != null) {
                    return _tableProperties.TableWidth.Type?.Value;
                }

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableWidth == null) {
                    if (value.HasValue) {
                        _tableProperties.TableWidth = new TableWidth() {
                            Type = value.Value,
                            Width = value.Value == TableWidthUnitValues.Auto ? "0" : "5000"
                        };
                    }
                } else {
                    if (value.HasValue) {
                        _tableProperties.TableWidth.Type = value.Value;
                        if (value.Value == TableWidthUnitValues.Auto) {
                            _tableProperties.TableWidth.Width = "0";
                        }
                    } else {
                        _tableProperties.TableWidth.Remove();
                    }
                }
                // Keep tblGrid consistent when width interpretation changes
                if (!_suppressGridRefresh) { try { RefreshTblGridFromColumnWidths(); } catch { } }
            }
        }

        /// <summary>
        /// Gets or sets width of a table
        /// </summary>
        public int? Width {
            get {
                if (_tableProperties != null && _tableProperties.TableWidth != null) {
                    if (!string.IsNullOrEmpty(_tableProperties.TableWidth.Width)) {
                        if (int.TryParse(_tableProperties.TableWidth.Width, out var width)) {
                            return width;
                        }
                    }
                }
                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableWidth == null) {
                    _tableProperties.TableWidth = new TableWidth() {
                        Type = TableWidthUnitValues.Pct,
                        Width = value?.ToString()
                    };
                } else {
                    _tableProperties.TableWidth.Width = value?.ToString();
                }
                // Grid depends on overall width (for pct columns)
                if (!_suppressGridRefresh) { try { RefreshTblGridFromColumnWidths(); } catch { } }
            }
        }

        /// <summary>
        /// Gets or sets layout of a table
        /// </summary>
        public TableLayoutValues? LayoutType {
            get {
                if (_tableProperties != null && _tableProperties.TableLayout != null) {
                    return _tableProperties.TableLayout.Type?.Value;
                }
                return TableLayoutValues.Autofit;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableLayout == null) {
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
                    return _tableProperties.TableLook.FirstRow?.Value;
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
                    return _tableProperties.TableLook.LastRow?.Value;
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
                    return _tableProperties.TableLook.FirstColumn?.Value;
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
                    return _tableProperties.TableLook.LastColumn?.Value;
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
                    return _tableProperties.TableLook.NoHorizontalBand?.Value;
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
                    return _tableProperties.TableLook.NoVerticalBand?.Value;
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

        /// <summary>
        /// Gets the number of rows in the table.
        /// </summary>
        public int RowsCount => this.Rows.Count;

        /// <summary>
        /// Gets the collection of rows belonging to the table.
        /// </summary>
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

        /// <summary>
        /// Gets the first row of the table.
        /// </summary>
        public WordTableRow FirstRow {
            get {
                return Rows.First();
            }
        }
        /// <summary>
        /// Gets the last row of the table.
        /// </summary>
        public WordTableRow LastRow {
            get {
                return Rows.Last();
            }
        }

        internal Table _table;

        internal TableProperties? _tableProperties {
            get {
                return _table.ChildElements.OfType<TableProperties>().FirstOrDefault();
            }
        }

        private WordDocument _document;
        //internal string Text;
        //private WordSection _section;

        private Header? _header {
            get {
                var parent = _table.Parent;
                if (parent is Header) {
                    return (Header)parent;
                }

                return null;
            }
        }

        private Footer? _footer {
            get {
                var parent = _table.Parent;
                if (parent is Footer) {
                    return (Footer)parent;
                }

                return null;
            }
        }

        /// <summary>
        /// Provides positioning information for the table within the document.
        /// </summary>
        public WordTablePosition Position;

        /// <summary>
        /// Gets the table style details. WIP
        /// </summary>
        public WordTableStyleDetails? StyleDetails {
            get {
                if (_tableProperties != null && _tableProperties.TableStyle != null) {
                    return new WordTableStyleDetails(this);
                }
                return null;
            }
        }


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
            // Ensure tblGrid mirrors initial cell widths so online viewers render correctly
            // (ColumnWidth defaults to DXA 2400 per cell at creation time).
            try { RefreshTblGridFromColumnWidths(); } catch { }
            return table;
        }

        // Prevents recursive RefreshTblGridFromColumnWidths() calls when setters adjust
        // table width/widthType during normalization.
        internal bool _suppressGridRefresh = false;

        /// <summary>
        /// Used during load of the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="table"></param>
        internal WordTable(WordDocument document, Table table) : this(document, table, initializeChildren: true) {
        }

        internal WordTable(WordDocument document, Table table, bool initializeChildren) {
            _table = table;
            _document = document;

            if (initializeChildren) {
                foreach (TableRow row in table.ChildElements.OfType<TableRow>().ToList()) {
                    _ = new WordTableRow(this, row, document);
                }
            }

            // Establish Position property
            Position = new WordTablePosition(this);
        }

        /// <summary>
        /// Creates a table instance without inserting it into the document.
        /// </summary>
        /// <param name="document">Parent <see cref="WordDocument"/>.</param>
        /// <param name="rows">Number of rows.</param>
        /// <param name="columns">Number of columns.</param>
        /// <param name="tableStyle">Style to apply to the table.</param>
        /// <returns>The newly created <see cref="WordTable"/>.</returns>
        public static WordTable Create(WordDocument document, int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            return new WordTable(document, rows, columns, tableStyle, insert: false);
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

            _document.InvalidateValidationCache();
        }

        /// <summary>
        /// Initializes a new instance of <see cref="WordTable"/> and optionally inserts it into the document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="rows">Number of rows.</param>
        /// <param name="columns">Number of columns.</param>
        /// <param name="tableStyle">Style applied to the table.</param>
        /// <param name="insert">If set to <c>true</c> the table is appended to the document immediately.</param>
        internal WordTable(WordDocument document, int rows, int columns, WordTableStyle tableStyle, bool insert = true) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            if (insert) {
                // Append the table to the document.
                document._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Append(_table);

                _document.InvalidateValidationCache();
            }
        }

        /// <summary>
        /// Initializes a new instance of <see cref="WordTable"/> inside the specified <see cref="TableCell"/>.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="wordTableCell">Table cell that will host the table.</param>
        /// <param name="rows">Number of rows.</param>
        /// <param name="columns">Number of columns.</param>
        /// <param name="tableStyle">Style applied to the table.</param>
        public WordTable(WordDocument document, TableCell wordTableCell, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;

            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            wordTableCell.Append(_table);

            _document.InvalidateValidationCache();
        }

        internal WordTable(WordDocument document, Footer footer, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            footer.Append(_table);

            _document.InvalidateValidationCache();
        }
        internal WordTable(WordDocument document, Header header, int rows, int columns, WordTableStyle tableStyle) {
            _document = document;
            _table = GenerateTable(document, rows, columns, tableStyle);

            // Establish Position property
            Position = new WordTablePosition(this);

            header.Append(_table);

            _document.InvalidateValidationCache();
        }

        /// <summary>
        /// Add row to an existing table with the specified number of columns
        /// </summary>
        /// <param name="cellsCount"></param>
        public WordTableRow AddRow(int cellsCount = 0) {
            WordTableRow row = new WordTableRow(_document, this);
            _table.Append(row._tableRow);
            AddCells(row, cellsCount);
            try { RefreshTblGridFromColumnWidths(); } catch { }
            _document.InvalidateValidationCache();
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
            try { RefreshTblGridFromColumnWidths(); } catch { }
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
            _document.InvalidateValidationCache();
            return rows;
        }

        /// <summary>
        /// Adds a row to the table using WordTableRow object
        /// </summary>
        /// <param name="row"></param>
        private void AddRow(WordTableRow row) {
            _table.Append(row._tableRow);
            _document.InvalidateValidationCache();
        }

        /// <summary>
        /// Remove table from document
        /// </summary>
        public void Remove() {
            _table.Remove();
            _document.InvalidateValidationCache();
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
