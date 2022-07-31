using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>
        ///     Gets or sets a Title/Caption to a Table
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
        ///     Gets or sets Description for a Table
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
        ///     Allow table to overlap or not
        /// </summary>
        public bool AllowOverlap {
            get {
                if (Position.TableOverlap == TableOverlapValues.Overlap) return true;
                return false;
            }
            set => Position.TableOverlap = value ? TableOverlapValues.Overlap : TableOverlapValues.Never;
        }

        /// <summary>
        ///     Allow text to wrap around table.
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
                for (int cellIndex = 0; cellIndex >= this.Rows[0].CellsCount; cellIndex++) {
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
    }
}