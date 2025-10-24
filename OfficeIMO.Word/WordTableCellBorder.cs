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
        // Do not force-create TableCellProperties on construction; keep read-only access non-mutating.
        // All writes ensure properties exist via helpers below.
        private TableCellProperties? TcPr => _wordTableCell._tableCellProperties;
        private TableCellProperties EnsurePr() {
            if (_wordTableCell._tableCellProperties == null) {
                _wordTableCell.AddTableCellProperties();
            }
            return _wordTableCell._tableCellProperties!;
        }
        private TableCellBorders? BordersOrNull => TcPr?.TableCellBorders;
        private TableCellBorders EnsureBorders() {
            var pr = EnsurePr();
            return pr.TableCellBorders ??= new TableCellBorders();
        }
        private LeftBorder EnsureLeft() { var b = EnsureBorders(); return b.LeftBorder ??= new LeftBorder(); }
        private RightBorder EnsureRight() { var b = EnsureBorders(); return b.RightBorder ??= new RightBorder(); }
        private TopBorder EnsureTop() { var b = EnsureBorders(); return b.TopBorder ??= new TopBorder(); }
        private BottomBorder EnsureBottom() { var b = EnsureBorders(); return b.BottomBorder ??= new BottomBorder(); }
        private InsideHorizontalBorder EnsureInsideHorizontal() { var b = EnsureBorders(); return b.InsideHorizontalBorder ??= new InsideHorizontalBorder(); }
        private InsideVerticalBorder EnsureInsideVertical() { var b = EnsureBorders(); return b.InsideVerticalBorder ??= new InsideVerticalBorder(); }
        private StartBorder EnsureStart() { var b = EnsureBorders(); return b.StartBorder ??= new StartBorder(); }
        private EndBorder EnsureEnd() { var b = EnsureBorders(); return b.EndBorder ??= new EndBorder(); }
        private TopLeftToBottomRightCellBorder EnsureTLBR() { var b = EnsureBorders(); return b.TopLeftToBottomRightCellBorder ??= new TopLeftToBottomRightCellBorder(); }
        private TopRightToBottomLeftCellBorder EnsureTRBL() { var b = EnsureBorders(); return b.TopRightToBottomLeftCellBorder ??= new TopRightToBottomLeftCellBorder(); }

        internal WordTableCellBorder(WordDocument wordDocument, WordTable wordTable, WordTableRow wordTableRow, WordTableCell wordTableCell) {
            _document = wordDocument;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _wordTableCell = wordTableCell;
            // Intentionally do not create TableCellProperties here.
        }

        /// <summary>
        /// Get or set left table cell border style
        /// </summary>
        public BorderValues? LeftStyle {
            get {
                return BordersOrNull?.LeftBorder?.Val?.Value;
            }
            set {
                var lb = EnsureLeft();
                lb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set left table cell border color using hex color codes
        /// </summary>
        public string? LeftColorHex {
            get {
                return BordersOrNull?.LeftBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var lb = EnsureLeft();
                lb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.LeftBorder?.Space;
            }
            set {
                var lb = EnsureLeft();
                lb.Space = value;
            }
        }

        /// <summary>
        /// Get or set left table cell border size
        /// </summary>
        public UInt32Value? LeftSize {
            get {
                return BordersOrNull?.LeftBorder?.Size;
            }
            set {
                var lb = EnsureLeft();
                lb.Size = value;
            }
        }


        /// <summary>
        /// Get or set right table cell border style
        /// </summary>
        public BorderValues? RightStyle {
            get {
                return BordersOrNull?.RightBorder?.Val?.Value;
            }
            set {
                var rb = EnsureRight();
                rb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set right table cell border color using hex color codes
        /// </summary>
        public string? RightColorHex {
            get {
                return BordersOrNull?.RightBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var rb = EnsureRight();
                rb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.RightBorder?.Space;
            }
            set {
                var rb = EnsureRight();
                rb.Space = value;
            }
        }

        /// <summary>
        /// Get or set right table cell border size
        /// </summary>
        public UInt32Value? RightSize {
            get {
                return BordersOrNull?.RightBorder?.Size;
            }
            set {
                var rb = EnsureRight();
                rb.Size = value;
            }
        }




        /// <summary>
        /// Get or set top table cell border style
        /// </summary>
        public BorderValues? TopStyle {
            get {
                return BordersOrNull?.TopBorder?.Val?.Value;
            }
            set {
                var tb = EnsureTop();
                tb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top table cell border color using hex color codes
        /// </summary>
        public string? TopColorHex {
            get {
                return BordersOrNull?.TopBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var tb = EnsureTop();
                tb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.TopBorder?.Space;
            }
            set {
                var tb = EnsureTop();
                tb.Space = value;
            }
        }

        /// <summary>
        /// Get or set top table cell border size
        /// </summary>
        public UInt32Value? TopSize {
            get {
                return BordersOrNull?.TopBorder?.Size;
            }
            set {
                var tb = EnsureTop();
                tb.Size = value;
            }
        }







        /// <summary>
        /// Get or set bottom table cell border style
        /// </summary>
        public BorderValues? BottomStyle {
            get {
                return BordersOrNull?.BottomBorder?.Val?.Value;
            }
            set {
                var bb = EnsureBottom();
                bb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using hex color codes
        /// </summary>
        public string? BottomColorHex {
            get {
                return BordersOrNull?.BottomBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var bb = EnsureBottom();
                bb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.BottomBorder?.Space;
            }
            set {
                var bb = EnsureBottom();
                bb.Space = value;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border size
        /// </summary>
        public UInt32Value? BottomSize {
            get {
                return BordersOrNull?.BottomBorder?.Size;
            }
            set {
                var bb = EnsureBottom();
                bb.Size = value;
            }
        }









        /// <summary>
        /// Get or set inside horizontal table cell border style
        /// </summary>
        public BorderValues? InsideHorizontalStyle {
            get {
                return BordersOrNull?.InsideHorizontalBorder?.Val?.Value;
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using hex color codes
        /// </summary>
        public string? InsideHorizontalColorHex {
            get {
                return BordersOrNull?.InsideHorizontalBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.InsideHorizontalBorder?.Space;
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Space = value;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border size
        /// </summary>
        public UInt32Value? InsideHorizontalSize {
            get {
                return BordersOrNull?.InsideHorizontalBorder?.Size;
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Size = value;
            }
        }









        /// <summary>
        /// Get or set inside vertical table cell border style
        /// </summary>
        public BorderValues? InsideVerticalStyle {
            get {
                return BordersOrNull?.InsideVerticalBorder?.Val?.Value;
            }
            set {
                var vb = EnsureInsideVertical();
                vb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using hex color codes
        /// </summary>
        public string? InsideVerticalColorHex {
            get {
                return BordersOrNull?.InsideVerticalBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var vb = EnsureInsideVertical();
                vb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.InsideVerticalBorder?.Space;
            }
                set {
                var vb = EnsureInsideVertical();
                vb.Space = value;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border size
        /// </summary>
        public UInt32Value? InsideVerticalSize {
            get {
                return BordersOrNull?.InsideVerticalBorder?.Size;
            }
            set {
                var vb = EnsureInsideVertical();
                vb.Size = value;
            }
        }






        /// <summary>
        /// Get or set start table cell border style
        /// </summary>
        public BorderValues? StartStyle {
            get {
                return BordersOrNull?.StartBorder?.Val?.Value;
            }
            set {
                var sb = EnsureStart();
                sb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set start table cell border color using hex color codes
        /// </summary>
        public string? StartColorHex {
            get {
                return BordersOrNull?.StartBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var sb = EnsureStart();
                sb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.StartBorder?.Space;
            }
            set {
                var sb = EnsureStart();
                sb.Space = value;
            }
        }

        /// <summary>
        /// Get or set start table cell border size
        /// </summary>
        public UInt32Value? StartSize {
            get {
                return BordersOrNull?.StartBorder?.Size;
            }
            set {
                var sb = EnsureStart();
                sb.Size = value;
            }
        }






        /// <summary>
        /// Get or set end table cell border style
        /// </summary>
        public BorderValues? EndStyle {
            get {
                return BordersOrNull?.EndBorder?.Val?.Value;
            }
            set {
                var eb = EnsureEnd();
                eb.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set end table cell border color using hex color codes
        /// </summary>
        public string? EndColorHex {
            get {
                return BordersOrNull?.EndBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var eb = EnsureEnd();
                eb.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.EndBorder?.Space;
            }
            set {
                var eb = EnsureEnd();
                eb.Space = value;
            }
        }

        /// <summary>
        /// Get or set end table cell border size
        /// </summary>
        public UInt32Value? EndSize {
            get {
                return BordersOrNull?.EndBorder?.Size;
            }
            set {
                var eb = EnsureEnd();
                eb.Size = value;
            }
        }













        /// <summary>
        /// Get or set top left to bottom right table cell border style
        /// </summary>
        public BorderValues? TopLeftToBottomRightStyle {
            get {
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Val?.Value;
            }
            set {
                var d = EnsureTLBR();
                d.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using hex color codes
        /// </summary>
        public string? TopLeftToBottomRightColorHex {
            get {
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var d = EnsureTLBR();
                d.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Space;
            }
            set {
                var d = EnsureTLBR();
                d.Space = value;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border size
        /// </summary>
        public UInt32Value? TopLeftToBottomRightSize {
            get {
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Size;
            }
            set {
                var d = EnsureTLBR();
                d.Size = value;
            }
        }




        /// <summary>
        /// Get or set top right to bottom left table cell border style
        /// </summary>
        public BorderValues? TopRightToBottomLeftStyle {
            get {
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Val?.Value;
            }
            set {
                var d = EnsureTRBL();
                d.Val = value.HasValue ? value.Value : null;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using hex color codes
        /// </summary>
        public string? TopRightToBottomLeftColorHex {
            get {
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Color?.Value?.Replace("#", "").ToLowerInvariant();
            }
            set {
                var d = EnsureTRBL();
                d.Color = value?.Replace("#", "").ToLowerInvariant();
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
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Space;
            }
            set {
                var d = EnsureTRBL();
                d.Space = value;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border size
        /// </summary>
        public UInt32Value? TopRightToBottomLeftSize {
            get {
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Size;
            }
            set {
                var d = EnsureTRBL();
                d.Size = value;
            }
        }

    }
}
