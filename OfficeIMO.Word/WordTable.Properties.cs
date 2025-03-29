using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>
        /// Gets or sets a Title/Caption to a Table
        /// </summary>
        public string Title {
            get {
                if (_tableProperties != null && _tableProperties.TableCaption != null)
                    return _tableProperties.TableCaption.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableCaption == null) _tableProperties.TableCaption = new TableCaption();
                if (value != null)
                    _tableProperties.TableCaption.Val = value;
                else
                    _tableProperties.TableCaption.Remove();
            }
        }

        /// <summary>
        /// Gets or sets Description for a Table
        /// </summary>
        public string Description {
            get {
                if (_tableProperties != null && _tableProperties.TableDescription != null)
                    return _tableProperties.TableDescription.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties.TableDescription == null)
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
        /// Sets or gets grid columns width (not really doing anything as far as I can see)
        /// </summary>
        public List<int> GridColumnWidth {
            get {
                var listReturn = new List<int>();
                TableGrid tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    var list = tableGrid.OfType<GridColumn>();
                    foreach (var column in list) {
                        listReturn.Add(int.Parse(column.Width.Value));
                    }
                }
                return listReturn;
            }
            set {
                TableGrid tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    tableGrid.RemoveAllChildren();
                } else {
                    _table.InsertAfter(new TableGrid(), _tableProperties);
                    tableGrid = _table.GetFirstChild<TableGrid>();
                }
                foreach (var columnWidth in value) {
                    tableGrid.Append(new GridColumn { Width = columnWidth.ToString() });
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
                    listReturn.Add(this.Rows[0].Cells[cellIndex].Width.Value);
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
        /// Get or Set Table Row Height for 1st row
        /// </summary>
        public List<int> RowHeight {
            get {
                var listReturn = new List<int>();
                // we assume the first row has the same widths as all rows, which may or may not be true
                for (int rowIndex = 0; rowIndex >= this.Rows.Count; rowIndex++) {
                    listReturn.Add(this.Rows[rowIndex].Height.Value);
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
        public WordTable ParentTable {
            get {
                if (IsNestedTable) {
                    Table table = (DocumentFormat.OpenXml.Wordprocessing.Table)this._table.Parent.Parent.Parent;
                    return new WordTable(this._document, table);
                }

                return null;
            }
        }
    }
}
