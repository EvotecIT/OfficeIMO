using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides a rich API for configuring the borders of a
    /// <see cref="WordTableCell"/>.  Every edge of the cell—including the
    /// diagonals—has properties for style, color, spacing and size so that
    /// callers can individually tailor each side of a cell.
    /// </summary>
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
            if (wordTableCell._tableCellProperties == null) {
                wordTableCell.AddTableCellProperties();
            }

            _tableCellProperties = wordTableCell._tableCellProperties!;
        }

        /// <summary>
        /// Get or set left table cell border style
        /// </summary>
        public BorderValues? LeftStyle {
            get {
                if (_tableCellProperties.TableCellBorders != null && _tableCellProperties.TableCellBorders.LeftBorder != null) {
                    return _tableCellProperties.TableCellBorders.LeftBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.LeftBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set left table cell border color using hex color codes
        /// </summary>
        public string? LeftColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.LeftBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.LeftBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set left table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color LeftColor {
            get {
                var hex = LeftColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("LeftColorHex is null"));
            }
            set { this.LeftColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set left table cell border space
        /// </summary>
        public UInt32Value? LeftSpace {
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
        public UInt32Value? LeftSize {
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
                    return _tableCellProperties.TableCellBorders.RightBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.RightBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set right table cell border color using hex color codes
        /// </summary>
        public string? RightColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.RightBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.RightBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set right table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color RightColor {
            get {
                var hex = RightColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("RightColorHex is null"));
            }
            set { this.RightColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set right table cell border space
        /// </summary>
        public UInt32Value? RightSpace {
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
        public UInt32Value? RightSize {
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
                    return _tableCellProperties.TableCellBorders.TopBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.TopBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top table cell border color using hex color codes
        /// </summary>
        public string? TopColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.TopBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.TopBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set top table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopColor {
            get {
                var hex = TopColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopColorHex is null"));
            }
            set { this.TopColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top table cell border space
        /// </summary>
        public UInt32Value? TopSpace {
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
        public UInt32Value? TopSize {
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
                    return _tableCellProperties.TableCellBorders.BottomBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.BottomBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using hex color codes
        /// </summary>
        public string? BottomColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.BottomBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.BottomBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color BottomColor {
            get {
                var hex = BottomColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("BottomColorHex is null"));
            }
            set { this.BottomColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set bottom table cell border space
        /// </summary>
        public UInt32Value? BottomSpace {
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
        public UInt32Value? BottomSize {
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
                    return _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using hex color codes
        /// </summary>
        public string? InsideHorizontalColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.InsideHorizontalBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.InsideHorizontalBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color InsideHorizontalColor {
            get {
                var hex = InsideHorizontalColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("InsideHorizontalColorHex is null"));
            }
            set { this.InsideHorizontalColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border space
        /// </summary>
        public UInt32Value? InsideHorizontalSpace {
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
        public UInt32Value? InsideHorizontalSize {
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
                    return _tableCellProperties.TableCellBorders.InsideVerticalBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using hex color codes
        /// </summary>
        public string? InsideVerticalColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.InsideVerticalBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.InsideVerticalBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color InsideVerticalColor {
            get {
                var hex = InsideVerticalColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("InsideVerticalColorHex is null"));
            }
            set { this.InsideVerticalColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set inside vertical table cell border space
        /// </summary>
        public UInt32Value? InsideVerticalSpace {
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
        public UInt32Value? InsideVerticalSize {
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
                    return _tableCellProperties.TableCellBorders.StartBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.StartBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set start table cell border color using hex color codes
        /// </summary>
        public string? StartColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.StartBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.StartBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set start table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color StartColor {
            get {
                var hex = StartColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("StartColorHex is null"));
            }
            set { this.StartColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set start table cell border space
        /// </summary>
        public UInt32Value? StartSpace {
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
        public UInt32Value? StartSize {
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
                    return _tableCellProperties.TableCellBorders.EndBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.EndBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set end table cell border color using hex color codes
        /// </summary>
        public string? EndColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.EndBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.EndBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set end table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color EndColor {
            get {
                var hex = EndColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("EndColorHex is null"));
            }
            set { this.EndColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set end table cell border space
        /// </summary>
        public UInt32Value? EndSpace {
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
        public UInt32Value? EndSize {
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
                    return _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using hex color codes
        /// </summary>
        public string? TopLeftToBottomRightColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.TopLeftToBottomRightCellBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.TopLeftToBottomRightCellBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopLeftToBottomRightColor {
            get {
                var hex = TopLeftToBottomRightColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopLeftToBottomRightColorHex is null"));
            }
            set { this.TopLeftToBottomRightColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border space
        /// </summary>
        public UInt32Value? TopLeftToBottomRightSpace {
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
        public UInt32Value? TopLeftToBottomRightSize {
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
                    return _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Val?.Value;
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

                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using hex color codes
        /// </summary>
        public string? TopRightToBottomLeftColorHex {
            get {
                if (_tableCellProperties.TableCellBorders?.TopRightToBottomLeftCellBorder?.Color?.Value is string color) {
                    return color.Replace("#", "").ToLowerInvariant();
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
                _tableCellProperties.TableCellBorders.TopRightToBottomLeftCellBorder.Color = value?.Replace("#", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using named colors
        /// </summary>
        public SixLabors.ImageSharp.Color TopRightToBottomLeftColor {
            get {
                var hex = TopRightToBottomLeftColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopRightToBottomLeftColorHex is null"));
            }
            set { this.TopRightToBottomLeftColorHex = value.ToHexColor(); }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border space
        /// </summary>
        public UInt32Value? TopRightToBottomLeftSpace {
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
        public UInt32Value? TopRightToBottomLeftSize {
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
