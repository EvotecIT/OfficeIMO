using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a single cell within a <see cref="WordTable"/>.
    /// </summary>
    public class WordTableCell : System.IEquatable<WordTableCell> {
        private WordTableCellBorder? _borders;

        /// <summary>
        /// Provides access to the border configuration of the cell.
        /// </summary>
        public WordTableCellBorder Borders => _borders ??= new WordTableCellBorder(_document, _wordTable, _wordTableRow, this);

        internal TableCell _tableCell;
        internal TableCellProperties? _tableCellProperties;

        /// <summary>
        /// Gets all <see cref="WordParagraph"/> instances contained in the cell.
        /// </summary>
        public List<WordParagraph> Paragraphs => WordSection.ConvertParagraphsToWordParagraphs(_document, _tableCell.ChildElements.OfType<Paragraph>());
        private readonly WordTable _wordTable;
        private readonly WordTableRow _wordTableRow;
        private readonly WordDocument _document;

        /// <summary>
        /// Gets the row that owns this cell.
        /// </summary>
        public WordTableRow Parent => _wordTableRow;

        /// <summary>
        /// Gets the table that owns this cell.
        /// </summary>
        public WordTable ParentTable => _wordTable;

        /// <summary>
        /// Gets or Sets Horizontal Merge for a Table Cell
        /// </summary>
        public MergedCellValues? HorizontalMerge {
            get {
                return _tableCellProperties?.HorizontalMerge?.Val?.Value;
            }
            set {
                AddTableCellProperties();
                if (value == null) {
                    _tableCellProperties!.HorizontalMerge?.Remove();
                } else {
                    _tableCellProperties!.HorizontalMerge ??= new HorizontalMerge();
                    _tableCellProperties.HorizontalMerge.Val = value.Value;
                }
            }
        }

        /// <summary>
        /// Gets or Sets Vertical Merge for a Table Cell
        /// </summary>
        public MergedCellValues? VerticalMerge {
            get {
                return _tableCellProperties?.VerticalMerge?.Val?.Value;
            }
            set {
                AddTableCellProperties();
                if (value == null) {
                    _tableCellProperties!.VerticalMerge?.Remove();
                } else {
                    _tableCellProperties!.VerticalMerge ??= new VerticalMerge();
                    _tableCellProperties.VerticalMerge.Val = value.Value;
                }
            }
        }

        /// <summary>
        /// Gets information whether the cell is part of a horizontal merge
        /// </summary>
        public bool HasHorizontalMerge {
            get {
                return _tableCellProperties?.HorizontalMerge != null;
            }
        }

        /// <summary>
        /// Gets information whether the cell is part of a vertical merge
        /// </summary>
        public bool HasVerticalMerge {
            get {
                return _tableCellProperties?.VerticalMerge != null;
            }
        }

        /// <summary>
        /// Get or set the background color of the cell using hexadecimal color code.
        /// </summary>
        public string ShadingFillColorHex {
            get {
                var fill = _tableCellProperties?.Shading?.Fill?.Value;
                return fill != null ? fill.ToLowerInvariant() : "";
            }
            set {
                AddTableCellProperties();
                if (value != "") {
                    var color = value.Replace("#", "").ToLowerInvariant();
                    _tableCellProperties!.Shading ??= new Shading();
                    _tableCellProperties.Shading.Fill = color;
                    _tableCellProperties.Shading.Val ??= ShadingPatternValues.Clear;
                } else {
                    if (_tableCellProperties?.Shading?.Fill != null) {
                        _tableCellProperties.Shading.Remove();
                    }
                }

            }
        }

        /// <summary>
        /// Add paragraph to the table cell
        /// </summary>
        /// <param name="paragraph">The paragraph to add to this cell, if
        /// this is not passed then a new empty paragraph with settings from
        /// the previous paragraph will be added.</param>
        /// <param name="removeExistingParagraphs">If value is not passed or false then add
        /// the given paragraph into the cell. If set to true then clear
        /// every existing paragraph before adding the new paragraph.
        /// </param>
        /// <returns>A reference to the added paragraph.</returns>
        public WordParagraph AddParagraph(WordParagraph? paragraph = null, bool removeExistingParagraphs = false) {
            // Considering between implementing a reset that clears all paragraphs or
            // a deletePrevious that will replace the last paragraph.
            // NOTE: Raise this during PR.
            if (removeExistingParagraphs) {
                var paragraphs = _tableCell.ChildElements.OfType<Paragraph>().ToList();
                foreach (var wordParagraph in paragraphs) {
                    wordParagraph.Remove();
                }
            }
            if (paragraph == null) {
                paragraph = new WordParagraph(this._document);
            }
            WordParagraph.EnsureParagraphCanBeInserted(this._document, _tableCell, paragraph,
                "append a paragraph to the table cell");
            _tableCell.Append(paragraph._paragraph);
            paragraph.RefreshParent();
            return paragraph;
        }

        /// <summary>
        /// Add paragraph to the table cell with text
        /// </summary>
        /// <param name="text"></param>
        /// <param name="removeExistingParagraphs"></param>
        /// <returns></returns>
        public WordParagraph AddParagraph(string text, bool removeExistingParagraphs = false) {
            return AddParagraph(paragraph: null, removeExistingParagraphs).SetText(text);
        }

        /// <summary>
        /// Get or set the background pattern of a cell
        /// </summary>
        public ShadingPatternValues? ShadingPattern {
            get {
                return _tableCellProperties?.Shading?.Val?.Value;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.Shading ??= new Shading();
                    _tableCellProperties.Shading.Val = value.Value;
                } else {
                    _tableCellProperties?.Shading?.Remove();
                }
            }
        }

        /// <summary>
        /// Get or set the background color of a cell using SixLabors.Color
        /// </summary>
        public Color? ShadingFillColor {
            get {
                if (ShadingFillColorHex != "") {
                    return Helpers.ParseColor(ShadingFillColorHex);
                }

                return null;
            }
            set {
                if (value != null) {
                    this.ShadingFillColorHex = value.Value.ToHexColor();
                }
            }
        }

        /// <summary>
        /// Gets or sets cell width
        /// </summary>
        public int? Width {
            get {
                var width = _tableCellProperties?.TableCellWidth?.Width;
                if (width != null) {
                    return int.Parse(width!);
                }

                return null;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellWidth ??= new TableCellWidth();
                    _tableCellProperties.TableCellWidth.Width = value.Value.ToString();
                } else {
                    _tableCellProperties?.TableCellWidth?.Remove();
                }
            }
        }

        //public int? WidthPercentage {
        //    get {
        //        if (_tableCellProperties.TableCellWidth != null) {
        //            var width = _tableCellProperties.TableCellWidth.Width;
        //            var type = _tableCellProperties.TableCellWidth.Type;
        //            if (type == TableWidthUnitValues.Pct) {
        //                if (width.Value.Contains("%")) {
        //                    return int.Parse(width.Value.Replace("%", ""));
        //                }
        //            } else if (type == TableWidthUnitValues.Dxa) {
        //                throw new NotImplementedException("WidthPercentage is not implemented for TableWidthUnitValues.Dxa");
        //                //var widthInInches = (double.Parse(width.Value) / 1440);
        //                //var widthInPercentage = (widthInInches / _wordTable.Width) * 100;
        //                //return (int)widthInPercentage;
        //            } else {
        //                throw new NotImplementedException("WidthPercentage is not implemented for " + type);
        //            }
        //        }

        //        return null;
        //    }
        //    set {

        //    }
        //}

        /// <summary>
        /// Gets or sets cell width type
        /// </summary>
        public TableWidthUnitValues? WidthType {
            get {
                return _tableCellProperties?.TableCellWidth?.Type?.Value;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellWidth ??= new TableCellWidth();
                    _tableCellProperties.TableCellWidth.Type = value.Value;
                } else {
                    _tableCellProperties?.TableCellWidth?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets text direction in a Table Cell
        /// </summary>
        public TextDirectionValues? TextDirection {
            get {
                return _tableCellProperties?.TextDirection?.Val?.Value;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TextDirection ??= new TextDirection();
                    _tableCellProperties.TextDirection.Val = value.Value;
                } else {
                    _tableCellProperties?.TextDirection?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets cell vertical alignment in a Table Cell
        /// </summary>
        public TableVerticalAlignmentValues? VerticalAlignment {
            get {
                return _tableCellProperties?.TableCellVerticalAlignment?.Val?.Value;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellVerticalAlignment ??= new TableCellVerticalAlignment();
                    _tableCellProperties.TableCellVerticalAlignment.Val = value.Value;
                } else {
                    _tableCellProperties?.TableCellVerticalAlignment?.Remove();
                }
            }
        }

        /// <summary>
        /// Removes the <see cref="TableCellMargin"/> element if it has no child elements.
        /// </summary>
        private void CleanupTableCellMargin() {
            if (_tableCellProperties?.TableCellMargin != null && !_tableCellProperties.TableCellMargin.Any()) {
                _tableCellProperties.TableCellMargin.Remove();
            }
        }

        /// <summary>
        /// Gets or sets the top margin in twips for the current cell.
        /// </summary>
        public Int16? MarginTopWidth {
            get {
                var width = _tableCellProperties?.TableCellMargin?.TopMargin?.Width;
                if (width != null) {
                    return short.Parse(width!);
                }
                return null;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellMargin ??= new TableCellMargin();
                    _tableCellProperties.TableCellMargin.TopMargin ??= new TopMargin();
                    _tableCellProperties.TableCellMargin.TopMargin.Width = value.ToString();
                    _tableCellProperties.TableCellMargin.TopMargin.Type = TableWidthUnitValues.Dxa;
                } else {
                    _tableCellProperties?.TableCellMargin?.TopMargin?.Remove();
                    CleanupTableCellMargin();
                }
            }
        }

        /// <summary>
        /// Gets or sets the bottom margin in twips for the current cell.
        /// </summary>
        public Int16? MarginBottomWidth {
            get {
                var width = _tableCellProperties?.TableCellMargin?.BottomMargin?.Width;
                if (width != null) {
                    return short.Parse(width!);
                }
                return null;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellMargin ??= new TableCellMargin();
                    _tableCellProperties.TableCellMargin.BottomMargin ??= new BottomMargin();
                    _tableCellProperties.TableCellMargin.BottomMargin.Width = value.ToString();
                    _tableCellProperties.TableCellMargin.BottomMargin.Type = TableWidthUnitValues.Dxa;
                } else {
                    _tableCellProperties?.TableCellMargin?.BottomMargin?.Remove();
                    CleanupTableCellMargin();
                }
            }
        }

        /// <summary>
        /// Gets or sets the left margin in twips for the current cell.
        /// </summary>
        public Int16? MarginLeftWidth {
            get {
                var width = _tableCellProperties?.TableCellMargin?.LeftMargin?.Width;
                if (width != null) {
                    return short.Parse(width!);
                }
                return null;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellMargin ??= new TableCellMargin();
                    _tableCellProperties.TableCellMargin.LeftMargin ??= new LeftMargin();
                    _tableCellProperties.TableCellMargin.LeftMargin.Width = value.ToString();
                    _tableCellProperties.TableCellMargin.LeftMargin.Type = TableWidthUnitValues.Dxa;
                } else {
                    _tableCellProperties?.TableCellMargin?.LeftMargin?.Remove();
                    CleanupTableCellMargin();
                }
            }
        }

        /// <summary>
        /// Gets or sets the right margin in twips for the current cell.
        /// </summary>
        public Int16? MarginRightWidth {
            get {
                var width = _tableCellProperties?.TableCellMargin?.RightMargin?.Width;
                if (width != null) {
                    return short.Parse(width!);
                }
                return null;
            }
            set {
                AddTableCellProperties();
                if (value != null) {
                    _tableCellProperties!.TableCellMargin ??= new TableCellMargin();
                    _tableCellProperties.TableCellMargin.RightMargin ??= new RightMargin();
                    _tableCellProperties.TableCellMargin.RightMargin.Width = value.ToString();
                    _tableCellProperties.TableCellMargin.RightMargin.Type = TableWidthUnitValues.Dxa;
                } else {
                    _tableCellProperties?.TableCellMargin?.RightMargin?.Remove();
                    CleanupTableCellMargin();
                }
            }
        }

        /// <summary>
        /// Gets or sets the top margin in centimeters for the current cell.
        /// </summary>
        public double? MarginTopCentimeters {
            get {
                if (MarginTopWidth != null) {
                    return Helpers.ConvertTwipsToCentimeters(MarginTopWidth.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    MarginTopWidth = (Int16)Helpers.ConvertCentimetersToTwips(value.Value);
                } else {
                    MarginTopWidth = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the bottom margin in centimeters for the current cell.
        /// </summary>
        public double? MarginBottomCentimeters {
            get {
                if (MarginBottomWidth != null) {
                    return Helpers.ConvertTwipsToCentimeters(MarginBottomWidth.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    MarginBottomWidth = (Int16)Helpers.ConvertCentimetersToTwips(value.Value);
                } else {
                    MarginBottomWidth = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the left margin in centimeters for the current cell.
        /// </summary>
        public double? MarginLeftCentimeters {
            get {
                if (MarginLeftWidth != null) {
                    return Helpers.ConvertTwipsToCentimeters(MarginLeftWidth.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    MarginLeftWidth = (Int16)Helpers.ConvertCentimetersToTwips(value.Value);
                } else {
                    MarginLeftWidth = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the right margin in centimeters for the current cell.
        /// </summary>
        public double? MarginRightCentimeters {
            get {
                if (MarginRightWidth != null) {
                    return Helpers.ConvertTwipsToCentimeters(MarginRightWidth.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    MarginRightWidth = (Int16)Helpers.ConvertCentimetersToTwips(value.Value);
                } else {
                    MarginRightWidth = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets whether text wraps within the cell.
        /// </summary>
        public bool WrapText {
            get {
                return _tableCellProperties?.GetFirstChild<NoWrap>() == null;
            }
            set {
                AddTableCellProperties();
                var current = _tableCellProperties!.GetFirstChild<NoWrap>();
                if (value) {
                    current?.Remove();
                } else {
                    if (current == null) {
                        _tableCellProperties.Append(new NoWrap());
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets whether text is compressed to fit within the cell width.
        /// </summary>
        public bool FitText {
            get {
                var tcPr = _tableCell.GetFirstChild<TableCellProperties>();
                return tcPr?.GetFirstChild<TableCellFitText>() != null;
            }
            set {
                AddTableCellProperties();
                var current = _tableCellProperties!.GetFirstChild<TableCellFitText>();
                if (value) {
                    if (current == null) {
                        _tableCellProperties.Append(new TableCellFitText { Val = OnOffOnlyValues.On });
                    } else {
                        current.Val = OnOffOnlyValues.On;
                    }
                } else {
                    current?.Remove();
                }
            }
        }


        /// <summary>
        /// Create a WordTableCell and add it to given Table Row
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wordTable"></param>
        /// <param name="wordTableRow"></param>
        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow) {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" });

            // Specify the width property of the table cell.
            tableCell.Append(tableCellProperties);

            WordParagraph paragraph = new WordParagraph();

            tableCell.Append(paragraph._paragraph);

            wordTableRow._tableRow.Append(tableCell);

            _tableCellProperties = tableCellProperties;
            _tableCell = tableCell;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;
        }

        /// <summary>
        /// Create a WordTableCell and add it to given Table Row from TableCell
        /// Mostly used for loading TableCells during Document Load
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wordTable"></param>
        /// <param name="wordTableRow"></param>
        /// <param name="tableCell"></param>
        /// <param name="ensureCellProperties">When true, provisions missing table cell properties to support editing.</param>
        internal WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow, TableCell tableCell, bool ensureCellProperties = true) {
            _tableCell = tableCell;
            if (ensureCellProperties && tableCell.TableCellProperties == null) {
                tableCell.TableCellProperties = new TableCellProperties();
            }

            _tableCellProperties = tableCell.TableCellProperties;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;

        }

        /// <summary>
        /// Ensure that TableCellProperties exist for current cell
        /// </summary>
        internal void AddTableCellProperties() {
            if (_tableCell.TableCellProperties == null) {
                _tableCell.InsertAt(new TableCellProperties(), 0);
            }
            _tableCellProperties = _tableCell.TableCellProperties!;
        }

        /// <summary>
        /// Remove TableCellProperties from current cell (used mostly for testing)
        /// </summary>
        internal void RemoveTableCellProperties() {
            if (_tableCell.TableCellProperties != null) {
                _tableCell.TableCellProperties.Remove();
                _tableCellProperties = null;
            }
        }

        /// <summary>
        /// Remove a cell from a table
        /// </summary>
        public void Remove() {
            _tableCell.Remove();
        }

        /// <summary>
        /// Merges two or more cells together horizontally.
        /// Provides ability to move or delete content of merged cells into single cell
        /// </summary>
        /// <param name="cellsCount"></param>
        /// <param name="copyParagraphs"></param>
        public void MergeHorizontally(int cellsCount, bool copyParagraphs = false) {
            var temporaryCell = _tableCell;
            AddTableCellProperties();
            _tableCellProperties!.HorizontalMerge = new HorizontalMerge {
                Val = MergedCellValues.Restart
            };

            for (int i = 0; i < cellsCount; i++) {
                var nextCell = _tableCell.NextSibling<TableCell>();
                if (nextCell != null) {
                    _tableCell = nextCell;
                    AddTableCellProperties();
                    if (copyParagraphs) {
                        var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                        foreach (var paragraph in paragraphs) {
                            paragraph.Remove();
                            temporaryCell.Append(paragraph);
                        }
                        _tableCell.Append(new Paragraph());
                    } else {
                        var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                        foreach (var paragraph in paragraphs) {
                            paragraph.Remove();
                        }
                        _tableCell.Append(new Paragraph());
                    }

                    _tableCell.TableCellProperties!.HorizontalMerge = new HorizontalMerge {
                        Val = MergedCellValues.Continue
                    };
                }
            }

        }

        /// <summary>
        /// Splits (unmerge) cells that were merged
        /// </summary>
        /// <param name="cellsCount"></param>
        public void SplitHorizontally(int cellsCount) {
            AddTableCellProperties();
            _tableCellProperties!.HorizontalMerge?.Remove();
            for (int i = 0; i < cellsCount; i++) {
                var nextCell = _tableCell.NextSibling<TableCell>();
                if (nextCell != null) {
                    _tableCell = nextCell;
                    AddTableCellProperties();
                    _tableCellProperties!.HorizontalMerge?.Remove();
                }
            }
        }

        /// <summary>
        /// Merges two or more cells together vertically
        /// </summary>
        /// <param name="cellsCount"></param>
        /// <param name="copyParagraphs"></param>
        public void MergeVertically(int cellsCount, bool copyParagraphs = false) {
            var temporaryCell = _tableCell;
            AddTableCellProperties();
            _tableCellProperties!.VerticalMerge = new VerticalMerge {
                Val = MergedCellValues.Restart
            };
            var tableRow = _tableCell.Parent as TableRow;
            var indexOfCell = tableRow?.ChildElements.ToList().IndexOf(_tableCell) ?? -1;

            for (int i = 0; i < cellsCount; i++) {
                if (tableRow != null) {
                    tableRow = tableRow.NextSibling<TableRow>();
                    if (tableRow != null && indexOfCell >= 0) {
                        var tableCells = tableRow.ChildElements.OfType<TableCell>().ToList();
                        if (indexOfCell < tableCells.Count) {
                            var nextCell = tableCells[indexOfCell];
                            _tableCell = nextCell;
                            AddTableCellProperties();
                            if (copyParagraphs) {
                                var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                                foreach (var paragraph in paragraphs) {
                                    paragraph.Remove();
                                    temporaryCell.Append(paragraph);
                                }
                                _tableCell.Append(new Paragraph());
                            } else {
                                var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                                foreach (var paragraph in paragraphs) {
                                    paragraph.Remove();
                                }
                                _tableCell.Append(new Paragraph());
                            }

                            _tableCellProperties!.VerticalMerge = new VerticalMerge {
                                Val = MergedCellValues.Continue
                            };
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Splits (unmerge) cells that were merged vertically
        /// </summary>
        /// <param name="cellsCount">Number of cells to split including the current one</param>
        public void SplitVertically(int cellsCount) {
            AddTableCellProperties();
            _tableCellProperties!.VerticalMerge?.Remove();

            var tableRow = _tableCell.Parent as TableRow;
            var indexOfCell = tableRow?.ChildElements.ToList().IndexOf(_tableCell) ?? -1;

            for (int i = 0; i < cellsCount; i++) {
                if (tableRow != null) {
                    tableRow = tableRow.NextSibling<TableRow>();
                    if (tableRow != null && indexOfCell >= 0) {
                        var tableCells = tableRow.ChildElements.OfType<TableCell>().ToList();
                        if (indexOfCell < tableCells.Count) {
                            _tableCell = tableCells[indexOfCell];
                            AddTableCellProperties();
                            _tableCellProperties!.VerticalMerge?.Remove();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Add table to a table cell (nested table)
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        /// <param name="tableStyle"></param>
        /// <param name="removePrecedingParagraph"></param>
        /// <returns></returns>
        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid, bool removePrecedingParagraph = false) {
            // Remove preceding empty paragraph by default (safe), or when explicitly requested.
            var paragraph = _tableCell.ChildElements.OfType<Paragraph>().LastOrDefault();
            if (paragraph != null) {
                bool hasText = paragraph.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text));
                if (removePrecedingParagraph || !hasText) {
                    paragraph.Remove();
                }
            }
            //this.Paragraphs[this.Paragraphs.Count - 1].Remove();
            WordTable wordTable = new WordTable(this._document, _tableCell, rows, columns, tableStyle);
            // Append a trailing empty paragraph (required by Word), but force zero spacing
            var trailing = new Paragraph();
            trailing.ParagraphProperties = trailing.ParagraphProperties ?? new ParagraphProperties();
            trailing.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() {
                Before = "0",
                After = "0",
                Line = "0"
            };
            _tableCell.Append(trailing);
            return wordTable;
        }

        /// <summary>
        /// Creates a list within this table cell using the specified style.
        /// </summary>
        /// <param name="style">List style to apply.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this.Paragraphs.Last());
            wordList.AddList(style);
            return wordList;
        }

        /// <summary>
        /// Gets information whether the cell contains other nested tables
        /// </summary>
        public bool HasNestedTables {
            get {
                return _tableCell.Descendants<Table>().Count() > 0;
            }
        }

        /// <summary>
        /// Get all nested tables in the cell
        /// </summary>
        public List<WordTable> NestedTables {
            get {
                var listReturn = new List<WordTable>();
                var list = _tableCell.Descendants<Table>().ToList();
                foreach (var table in list) {
                    listReturn.Add(new WordTable(this._document, table));
                }
                return listReturn;
            }
        }

        /// <summary>
        /// Determines whether this instance and another cell reference the same underlying OpenXML cell.
        /// </summary>
        public bool Equals(WordTableCell? other) {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            return ReferenceEquals(_tableCell, other._tableCell);
        }

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is WordTableCell other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => _tableCell?.GetHashCode() ?? 0;

        /// <summary>
        /// Compares two cells for equality based on the underlying OpenXML cell reference.
        /// </summary>
        public static bool operator ==(WordTableCell? left, WordTableCell? right) {
            if (left is null) return right is null;
            return left.Equals(right);
        }

        /// <summary>
        /// Determines whether two cells are not equal.
        /// </summary>
        public static bool operator !=(WordTableCell? left, WordTableCell? right) => !(left == right);
    }
}
