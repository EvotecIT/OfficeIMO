using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a table in a Word document and exposes various
    /// properties controlling its appearance and behavior.
    /// </summary>
    public partial class WordTable {
        /// <summary>
        /// Gets or sets a Title/Caption to a Table
        /// </summary>
        public string? Title {
            get {
                if (_tableProperties != null && _tableProperties.TableCaption != null)
                    return _tableProperties.TableCaption.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableCaption == null) _tableProperties.TableCaption = new TableCaption();
                if (value != null)
                    _tableProperties.TableCaption.Val = value;
                else
                    _tableProperties.TableCaption.Remove();
            }
        }

        /// <summary>
        /// Gets or sets Description for a Table
        /// </summary>
        public string? Description {
            get {
                if (_tableProperties != null && _tableProperties.TableDescription != null)
                    return _tableProperties.TableDescription.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableDescription == null)
                    _tableProperties.TableDescription = new TableDescription();
                if (value != null)
                    _tableProperties.TableDescription.Val = value;
                else
                    _tableProperties.TableDescription.Remove();
            }
        }

        /// <summary>
        /// Allow table to overlap or not
        /// </summary>
        public bool AllowOverlap {
            get {
                if (Position.TableOverlap == TableOverlapValues.Overlap) return true;
                return false;
            }
            set => Position.TableOverlap = value ? TableOverlapValues.Overlap : TableOverlapValues.Never;
        }

        /// <summary>
        /// Gets or sets the effective layout mode of the table using WordTableLayoutType enum.
        /// Setting FixedWidth via this property defaults to 100% width.
        /// Use SetFixedWidth(percentage) for specific percentages.
        /// </summary>
        public WordTableLayoutType LayoutMode {
            get => GetCurrentLayoutType();
            set {
                switch (value) {
                    case WordTableLayoutType.AutoFitToContents:
                        AutoFitToContents();
                        break;
                    case WordTableLayoutType.AutoFitToWindow:
                        AutoFitToWindow();
                        break;
                    case WordTableLayoutType.FixedWidth:
                        // Default to 100% when setting via this property
                        SetFixedWidth(100);
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets the AutoFit behavior for the table. Alias for <see cref="LayoutMode"/>.
        /// </summary>
        public WordTableLayoutType AutoFit {
            get => LayoutMode;
            set => LayoutMode = value;
        }

        /// <summary>
        /// Allow text to wrap around table.
        /// </summary>
        public bool AllowTextWrap {
            get {
                if (Position.VerticalAnchor == VerticalAnchorValues.Text) return true;

                return false;
            }
            set {
                if (value)
                    Position.VerticalAnchor = VerticalAnchorValues.Text;
                else
                    Position.VerticalAnchor = null;
            }
        }

        /// <summary>
        /// Gets or sets whether text wraps within all cells of the table.
        /// </summary>
        public bool WrapText {
            get {
                return Rows.SelectMany(row => row.Cells).All(cell => cell.WrapText);
            }
            set {
                foreach (var row in Rows) {
                    foreach (var cell in row.Cells) {
                        cell.WrapText = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets whether text is compressed to fit within all cells of the table.
        /// </summary>
        public bool FitText {
            get {
                return Rows.SelectMany(row => row.Cells).All(cell => cell.FitText);
            }
            set {
                foreach (var row in Rows) {
                    foreach (var cell in row.Cells) {
                        cell.FitText = value;
                    }
                }
            }
        }

        /// <summary>
        /// Sets or gets grid columns width (not really doing anything as far as I can see)
        /// </summary>
        public List<int> GridColumnWidth {
            get {
                var listReturn = new List<int>();
                TableGrid? tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    var list = tableGrid.OfType<GridColumn>();
                    foreach (var column in list) {
                        if (column.Width != null && column.Width.Value != null) {
                            listReturn.Add(int.Parse(column.Width.Value));
                        }
                    }
                }
                return listReturn;
            }
            set {
                TableGrid? tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    tableGrid.RemoveAllChildren();
                } else {
                    _table.InsertAfter(new TableGrid(), _tableProperties);
                    tableGrid = _table.GetFirstChild<TableGrid>();
                }
                if (tableGrid != null) {
                    foreach (var columnWidth in value) {
                        tableGrid.Append(new GridColumn { Width = columnWidth.ToString() });
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets column width for a whole table simplifying setup of column width
        /// Please note that it assumes first row has the same width as the rest of rows
        /// which may give false positives if there are multiple values set differently.
        /// </summary>
        public List<int> ColumnWidth {
            get {
                var listReturn = new List<int>();
                // we assume the first row has the same widths as all rows, which may or may not be true
                for (int cellIndex = 0; cellIndex < this.Rows[0].CellsCount; cellIndex++) {
                    var width = this.Rows[0].Cells[cellIndex].Width;
                    if (width.HasValue) {
                        listReturn.Add(width.Value);
                    }
                }
                return listReturn;
            }
            set {
                for (int cellIndex = 0; cellIndex < value.Count; cellIndex++) {
                    foreach (var row in this.Rows) {
                        row.Cells[cellIndex].Width = value[cellIndex];
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the column width type for a whole table simplifying setup of column width
        /// </summary>
        public TableWidthUnitValues? ColumnWidthType {
            get {
                var listReturn = new List<TableWidthUnitValues?>();
                // we assume the first row has the same widths as all rows, which may or may not be true
                for (int cellIndex = 0; cellIndex < this.Rows[0].CellsCount; cellIndex++) {
                    listReturn.Add(this.Rows[0].Cells[cellIndex].WidthType);
                }
                // we assume all cells have the same width type, which may or may not be true
                return listReturn[0];
            }
            set {
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        cell.WidthType = value;
                    }
                }
            }
        }

        /// <summary>
        /// Get or set row heights for the table
        /// </summary>
        public List<int> RowHeight {
            get {
                var listReturn = new List<int>();
                for (int rowIndex = 0; rowIndex < this.Rows.Count; rowIndex++) {
                    listReturn.Add(this.Rows[rowIndex].Height ?? 0);
                }
                return listReturn;
            }
            set {
                for (int rowIndex = 0; rowIndex < value.Count; rowIndex++) {
                    this.Rows[rowIndex].Height = value[rowIndex];
                }
            }

        }

        /// <summary>
        /// Get all WordTableCells in a table. A short way to loop thru all cells
        /// </summary>
        public List<WordTableCell> Cells {
            get {
                var listReturn = new List<WordTableCell>();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        listReturn.Add(cell);
                    }
                }
                return listReturn;
            }
        }


        /// <summary>
        /// Gets information whether the Table has other nested tables in at least one of the TableCells
        /// </summary>
        public bool HasNestedTables {
            get {
                foreach (var cell in this.Cells) {
                    if (cell.HasNestedTables) {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Get all nested tables in the table
        /// </summary>
        public List<WordTable> NestedTables {
            get {
                var listReturn = new List<WordTable>();
                foreach (var cell in this.Cells) {
                    listReturn.AddRange(cell.NestedTables);
                }
                return listReturn;
            }
        }

        /// <summary>
        /// Gets information whether the table is nested table (within TableCell)
        /// </summary>
        public bool IsNestedTable {
            get {
                var openXmlElement = this._table.Parent;
                if (openXmlElement != null) {
                    var typeOfParent = openXmlElement.GetType();
                    if (typeOfParent.FullName == "DocumentFormat.OpenXml.Wordprocessing.TableCell") {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Gets nested table parent table if table is nested table
        /// </summary>
        public WordTable? ParentTable {
            get {
                if (IsNestedTable) {
                    if (this._table.Parent?.Parent?.Parent is Table table) {
                        return new WordTable(this._document, table);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets all structured document tags contained in the table.
        /// </summary>
        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsStructuredDocumentTag).ToList();
                        foreach (var paragraph in paragraphs) {
                            list.Add(paragraph.StructuredDocumentTag);
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all checkbox content controls contained in the table.
        /// </summary>
        public List<WordCheckBox> CheckBoxes {
            get {
                List<WordCheckBox> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsCheckBox).ToList();
                        foreach (var paragraph in paragraphs) {
                            list.Add(paragraph.CheckBox);
                        }
                    }
                }
                return list;
            }
        }
        /// <summary>
        /// Gets all date picker content controls contained in the table.
        /// </summary>
        public List<WordDatePicker> DatePickers {
            get {
                List<WordDatePicker> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsDatePicker).ToList();
                        foreach (var paragraph in paragraphs) {
                            list.Add(paragraph.DatePicker);
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all dropdown list content controls contained in the table.
        /// </summary>
        public List<WordDropDownList> DropDownLists {
            get {
                List<WordDropDownList> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsDropDownList).ToList();
                        foreach (var paragraph in paragraphs) {
                            list.Add(paragraph.DropDownList);
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all repeating section content controls contained in the table.
        /// </summary>
        public List<WordRepeatingSection> RepeatingSections {
            get {
                List<WordRepeatingSection> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsRepeatingSection).ToList();
                        foreach (var paragraph in paragraphs) {
                            list.Add(paragraph.RepeatingSection);
                        }
                    }
                }
                return list;
            }
        }
    }
}