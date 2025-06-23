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
            get { return Helpers.ParseColor(LeftColorHex); }
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


        /// <summary>
        /// Get or set right table cell border style
        /// </summary>
        public BorderValues? RightStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.RightBorder != null) {
                    return _tableCellProperties.TableCellBorders.RightBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.RightBorder == null) {
                    _tableCellProperties.TableCellBorders.RightBorder = new RightBorder();
                }
                _tableCellProperties.TableCellBorders.RightBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set right table cell border color using hex color codes
        /// </summary>
        public string RightColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.RightBorder != null
                    && _tableCellProperties.TableCellBorders.RightBorder.Color != null
                    && _tableCellProperties.TableCellBorders.RightBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.RightBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.RightBorder == null) {
                    _tableCellProperties.TableCellBorders.RightBorder = new RightBorder();
                }
                _tableCellProperties.TableCellBorders.RightBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set right table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color RightColor {
            get { return Helpers.ParseColor(RightColorHex); }
            set { this.RightColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set right table cell border space
        /// </summary>
        public UInt32Value RightSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.RightBorder != null
                    && _tableCellProperties.TableCellBorders.RightBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.RightBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.RightBorder == null) {
                    _tableCellProperties.TableCellBorders.RightBorder = new RightBorder();
                }

                _tableCellProperties.TableCellBorders.RightBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set right table cell border size
        /// </summary>
        public UInt32Value RightSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.RightBorder != null
                    && _tableCellProperties.TableCellBorders.RightBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.RightBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.RightBorder == null) {
                    _tableCellProperties.TableCellBorders.RightBorder = new RightBorder();
                }

                _tableCellProperties.TableCellBorders.RightBorder.Size = value;
            }
        }




        /// <summary>
        /// Get or set top table cell border style
        /// </summary>
        public BorderValues? TopStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.TopBorder != null) {
                    return _tableCellProperties.TableCellBorders.TopBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopBorder == null) {
                    _tableCellProperties.TableCellBorders.TopBorder = new TopBorder();
                }
                _tableCellProperties.TableCellBorders.TopBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set top table cell border color using hex color codes
        /// </summary>
        public string TopColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopBorder != null
                    && _tableCellProperties.TableCellBorders.TopBorder.Color != null
                    && _tableCellProperties.TableCellBorders.TopBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.TopBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopBorder == null) {
                    _tableCellProperties.TableCellBorders.TopBorder = new TopBorder();
                }
                _tableCellProperties.TableCellBorders.TopBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set top table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopColor {
            get { return Helpers.ParseColor(TopColorHex); }
            set { this.TopColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top table cell border space
        /// </summary>
        public UInt32Value TopSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopBorder != null
                    && _tableCellProperties.TableCellBorders.TopBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.TopBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopBorder == null) {
                    _tableCellProperties.TableCellBorders.TopBorder = new TopBorder();
                }

                _tableCellProperties.TableCellBorders.TopBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set top table cell border size
        /// </summary>
        public UInt32Value TopSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopBorder != null
                    && _tableCellProperties.TableCellBorders.TopBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.TopBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopBorder == null) {
                    _tableCellProperties.TableCellBorders.TopBorder = new TopBorder();
                }

                _tableCellProperties.TableCellBorders.TopBorder.Size = value;
            }
        }







        /// <summary>
        /// Get or set bottom table cell border style
        /// </summary>
        public BorderValues? BottomStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.BottomBorder != null) {
                    return _tableCellProperties.TableCellBorders.BottomBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.BottomBorder == null) {
                    _tableCellProperties.TableCellBorders.BottomBorder = new BottomBorder();
                }
                _tableCellProperties.TableCellBorders.BottomBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using hex color codes
        /// </summary>
        public string BottomColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.BottomBorder != null
                    && _tableCellProperties.TableCellBorders.BottomBorder.Color != null
                    && _tableCellProperties.TableCellBorders.BottomBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.BottomBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.BottomBorder == null) {
                    _tableCellProperties.TableCellBorders.BottomBorder = new BottomBorder();
                }
                _tableCellProperties.TableCellBorders.BottomBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color BottomColor {
            get { return Helpers.ParseColor(BottomColorHex); }
            set { this.BottomColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set bottom table cell border space
        /// </summary>
        public UInt32Value BottomSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.BottomBorder != null
                    && _tableCellProperties.TableCellBorders.BottomBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.BottomBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.BottomBorder == null) {
                    _tableCellProperties.TableCellBorders.BottomBorder = new BottomBorder();
                }

                _tableCellProperties.TableCellBorders.BottomBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border size
        /// </summary>
        public UInt32Value BottomSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.BottomBorder != null
                    && _tableCellProperties.TableCellBorders.BottomBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.BottomBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.BottomBorder == null) {
                    _tableCellProperties.TableCellBorders.BottomBorder = new BottomBorder();
                }

                _tableCellProperties.TableCellBorders.BottomBorder.Size = value;
            }
        }









        /// <summary>
        /// Get or set inside horizontal table cell border style
        /// </summary>
        public BorderValues? InsideHorizontalStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.InsideHorizontalBorder != null) {
                    return _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideHorizontalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
                }
                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using hex color codes
        /// </summary>
        public string InsideHorizontalColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Color != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideHorizontalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
                }
                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color InsideHorizontalColor {
            get { return Helpers.ParseColor(InsideHorizontalColorHex); }
            set { this.InsideHorizontalColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border space
        /// </summary>
        public UInt32Value InsideHorizontalSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideHorizontalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
                }

                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border size
        /// </summary>
        public UInt32Value InsideHorizontalSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideHorizontalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
                }

                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Size = value;
            }
        }









        /// <summary>
        /// Get or set inside vertical table cell border style
        /// </summary>
        public BorderValues? InsideVerticalStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.InsideVerticalBorder != null) {
                    return _tableCellProperties.TableCellBorders.InsideVerticalBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideVerticalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideVerticalBorder = new InsideVerticalBorder();
                }
                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using hex color codes
        /// </summary>
        public string InsideVerticalColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder.Color != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.InsideVerticalBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideVerticalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideVerticalBorder = new InsideVerticalBorder();
                }
                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color InsideVerticalColor {
            get { return Helpers.ParseColor(InsideVerticalColorHex); }
            set { this.InsideVerticalColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set inside vertical table cell border space
        /// </summary>
        public UInt32Value InsideVerticalSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.InsideVerticalBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideVerticalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideVerticalBorder = new InsideVerticalBorder();
                }

                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border size
        /// </summary>
        public UInt32Value InsideVerticalSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder != null
                    && _tableCellProperties.TableCellBorders.InsideVerticalBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.InsideVerticalBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.InsideVerticalBorder == null) {
                    _tableCellProperties.TableCellBorders.InsideVerticalBorder = new InsideVerticalBorder();
                }

                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Size = value;
            }
        }






        /// <summary>
        /// Get or set start table cell border style
        /// </summary>
        public BorderValues? StartStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.StartBorder != null) {
                    return _tableCellProperties.TableCellBorders.StartBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.StartBorder == null) {
                    _tableCellProperties.TableCellBorders.StartBorder = new StartBorder();
                }
                _tableCellProperties.TableCellBorders.StartBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set start table cell border color using hex color codes
        /// </summary>
        public string StartColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.StartBorder != null
                    && _tableCellProperties.TableCellBorders.StartBorder.Color != null
                    && _tableCellProperties.TableCellBorders.StartBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.StartBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.StartBorder == null) {
                    _tableCellProperties.TableCellBorders.StartBorder = new StartBorder();
                }
                _tableCellProperties.TableCellBorders.StartBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set start table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color StartColor {
            get { return Helpers.ParseColor(StartColorHex); }
            set { this.StartColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set start table cell border space
        /// </summary>
        public UInt32Value StartSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.StartBorder != null
                    && _tableCellProperties.TableCellBorders.StartBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.StartBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.StartBorder == null) {
                    _tableCellProperties.TableCellBorders.StartBorder = new StartBorder();
                }

                _tableCellProperties.TableCellBorders.StartBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set start table cell border size
        /// </summary>
        public UInt32Value StartSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.StartBorder != null
                    && _tableCellProperties.TableCellBorders.StartBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.StartBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.StartBorder == null) {
                    _tableCellProperties.TableCellBorders.StartBorder = new StartBorder();
                }

                _tableCellProperties.TableCellBorders.StartBorder.Size = value;
            }
        }






        /// <summary>
        /// Get or set end table cell border style
        /// </summary>
        public BorderValues? EndStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.EndBorder != null) {
                    return _tableCellProperties.TableCellBorders.EndBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.EndBorder == null) {
                    _tableCellProperties.TableCellBorders.EndBorder = new EndBorder();
                }
                _tableCellProperties.TableCellBorders.EndBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set end table cell border color using hex color codes
        /// </summary>
        public string EndColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.EndBorder != null
                    && _tableCellProperties.TableCellBorders.EndBorder.Color != null
                    && _tableCellProperties.TableCellBorders.EndBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.EndBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.EndBorder == null) {
                    _tableCellProperties.TableCellBorders.EndBorder = new EndBorder();
                }
                _tableCellProperties.TableCellBorders.EndBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set end table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color EndColor {
            get { return Helpers.ParseColor(EndColorHex); }
            set { this.EndColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set end table cell border space
        /// </summary>
        public UInt32Value EndSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.EndBorder != null
                    && _tableCellProperties.TableCellBorders.EndBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.EndBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.EndBorder == null) {
                    _tableCellProperties.TableCellBorders.EndBorder = new EndBorder();
                }

                _tableCellProperties.TableCellBorders.EndBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set end table cell border size
        /// </summary>
        public UInt32Value EndSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.EndBorder != null
                    && _tableCellProperties.TableCellBorders.EndBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.EndBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.EndBorder == null) {
                    _tableCellProperties.TableCellBorders.EndBorder = new EndBorder();
                }

                _tableCellProperties.TableCellBorders.EndBorder.Size = value;
            }
        }













        /// <summary>
        /// Get or set top left to bottom right table cell border style
        /// </summary>
        public BorderValues? TopLeftToBottomRightStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder != null) {
                    return _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder();
                }
                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using hex color codes
        /// </summary>
        public string TopLeftToBottomRightColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Color != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder();
                }
                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopLeftToBottomRightColor {
            get { return Helpers.ParseColor(TopLeftToBottomRightColorHex); }
            set { this.TopLeftToBottomRightColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border space
        /// </summary>
        public UInt32Value TopLeftToBottomRightSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder();
                }

                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border size
        /// </summary>
        public UInt32Value TopLeftToBottomRightSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder();
                }

                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Size = value;
            }
        }




        /// <summary>
        /// Get or set top right to bottom left table cell border style
        /// </summary>
        public BorderValues? TopRightToBottomLeftStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder != null) {
                    return _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Val;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder();
                }
                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Val = value;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using hex color codes
        /// </summary>
        public string TopRightToBottomLeftColorHex {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Color != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Color.Value != null) {
                    return _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Color.Value.Replace("#", "");
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder();
                }
                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Color = value.Replace("#", "");
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopRightToBottomLeftColor {
            get { return Helpers.ParseColor(TopRightToBottomLeftColorHex); }
            set { this.TopRightToBottomLeftColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border space
        /// </summary>
        public UInt32Value TopRightToBottomLeftSpace {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Space != null) {
                    return _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Space;
                }
                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder();
                }

                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Space = value;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border size
        /// </summary>
        public UInt32Value TopRightToBottomLeftSize {
            get {
                if (_tableCellProperties.TableCellBorders != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder != null
                    && _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Size != null) {
                    return _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Size;
                }

                return null;
            }
            set {
                if (_tableCellProperties.TableCellBorders == null) {
                    _tableCellProperties.TableCellBorders = new TableCellBorders();
                }

                if (_tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder == null) {
                    _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder();
                }

                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Size = value;
            }
        }

    }
}
