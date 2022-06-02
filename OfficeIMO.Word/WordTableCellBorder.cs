using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableCellBorder {
        private readonly WordTableCell _wordTableCell;
        private readonly WordTableRow _wordTableRow;
        private readonly WordTable _wordTable;
        private readonly WordDocument _document;
        private readonly TableCellProperties _tableCellProperties;

        internal WordTableCellBorder(WordDocument wordDocument, WordTable wordTable, WordTableRow wordTableRow, WordTableCell wordTableCell) {
            _document = wordDocument;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _wordTableCell = wordTableCell;
            _tableCellProperties = wordTableCell._tableCellProperties;
        }

        /// <summary>
        /// Get or set left table cell border style
        /// </summary>
        public BorderValues? LeftStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.LeftBorder != null) {
                    return _tableCellProperties.TableCellBorders.LeftBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.LeftBorder == null) {
                    _tableCellProperties.TableCellBorders.LeftBorder = new LeftBorder();
                }
                _tableCellProperties.TableCellBorders.LeftBorder.Val = value;
            }
        }


        /// <summary>
        /// Get or set left table cell border color using hex color codes
        /// </summary>
        public string LeftColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.LeftBorder != null
                    && _tableCellProperties.TableCellBorders.LeftBorder.Color != null
                    && _tableCellProperties.TableCellBorders.LeftBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.LeftBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.LeftBorder == null) {
                    _tableCellProperties.TableCellBorders.LeftBorder = new LeftBorder();
                }
                _tableCellProperties.TableCellBorders.LeftBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set left table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color LeftColor {
            get { return SixLabors.ImageSharp.Color.Parse("#" + LeftColorHex); }
            set { this.LeftColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set left table cell border space
        /// </summary>
        public UInt32Value LeftSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.LeftBorder != null
                    && _tableCellProperties.TableCellBorders.LeftBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.LeftBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.LeftBorder == null) {
                    _tableCellProperties.TableCellBorders.LeftBorder = new LeftBorder();
                }

                _tableCellProperties.TableCellBorders.LeftBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set left table cell border size
        /// </summary>
        public UInt32Value LeftSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.LeftBorder != null
                    && _tableCellProperties.TableCellBorders.LeftBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.LeftBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.LeftBorder == null) {
                    _tableCellProperties.TableCellBorders.LeftBorder = new LeftBorder();
                }

                _tableCellProperties.TableCellBorders.LeftBorder.Size = value;
            }
        }
    }
}
