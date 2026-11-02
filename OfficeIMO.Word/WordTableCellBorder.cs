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
            if (pr.TableCellBorders == null) {
                pr.TableCellBorders = new TableCellBorders();
                _wordTableCell.NormalizeTableCellPropertiesOrder();
            }

            return pr.TableCellBorders;
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
        public WordBorderStyle? LeftStyle {
            get {
                return BordersOrNull?.LeftBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var lb = EnsureLeft();
                lb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set left table cell border color using hex color codes
        /// </summary>
        public string? LeftColorHex {
            get {
                return BordersOrNull?.LeftBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var lb = EnsureLeft();
                lb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set left table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor LeftColor {
            get {
                var hex = LeftColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("LeftColorHex is null"));
            }
            set { this.LeftColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the left cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.LeftBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var lb = EnsureLeft();
                lb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the left cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.LeftBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var lb = EnsureLeft();
                lb.Size = nativeValue;
            }
        }


        /// <summary>
        /// Get or set right table cell border style
        /// </summary>
        public WordBorderStyle? RightStyle {
            get {
                return BordersOrNull?.RightBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var rb = EnsureRight();
                rb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set right table cell border color using hex color codes
        /// </summary>
        public string? RightColorHex {
            get {
                return BordersOrNull?.RightBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var rb = EnsureRight();
                rb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set right table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor RightColor {
            get {
                var hex = RightColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("RightColorHex is null"));
            }
            set { this.RightColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the right cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.RightBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var rb = EnsureRight();
                rb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the right cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.RightBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var rb = EnsureRight();
                rb.Size = nativeValue;
            }
        }




        /// <summary>
        /// Get or set top table cell border style
        /// </summary>
        public WordBorderStyle? TopStyle {
            get {
                return BordersOrNull?.TopBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var tb = EnsureTop();
                tb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set top table cell border color using hex color codes
        /// </summary>
        public string? TopColorHex {
            get {
                return BordersOrNull?.TopBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var tb = EnsureTop();
                tb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set top table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor TopColor {
            get {
                var hex = TopColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopColorHex is null"));
            }
            set { this.TopColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the top cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.TopBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var tb = EnsureTop();
                tb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the top cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.TopBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var tb = EnsureTop();
                tb.Size = nativeValue;
            }
        }







        /// <summary>
        /// Get or set bottom table cell border style
        /// </summary>
        public WordBorderStyle? BottomStyle {
            get {
                return BordersOrNull?.BottomBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var bb = EnsureBottom();
                bb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using hex color codes
        /// </summary>
        public string? BottomColorHex {
            get {
                return BordersOrNull?.BottomBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var bb = EnsureBottom();
                bb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set bottom table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor BottomColor {
            get {
                var hex = BottomColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("BottomColorHex is null"));
            }
            set { this.BottomColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the bottom cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.BottomBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var bb = EnsureBottom();
                bb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the bottom cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.BottomBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var bb = EnsureBottom();
                bb.Size = nativeValue;
            }
        }









        /// <summary>
        /// Get or set inside horizontal table cell border style
        /// </summary>
        public WordBorderStyle? InsideHorizontalStyle {
            get {
                return BordersOrNull?.InsideHorizontalBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using hex color codes
        /// </summary>
        public string? InsideHorizontalColorHex {
            get {
                return BordersOrNull?.InsideHorizontalBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var hb = EnsureInsideHorizontal();
                hb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set inside horizontal table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor InsideHorizontalColor {
            get {
                var hex = InsideHorizontalColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("InsideHorizontalColorHex is null"));
            }
            set { this.InsideHorizontalColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the inside horizontal cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideHorizontalSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.InsideHorizontalBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var hb = EnsureInsideHorizontal();
                hb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the inside horizontal cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideHorizontalSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.InsideHorizontalBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var hb = EnsureInsideHorizontal();
                hb.Size = nativeValue;
            }
        }









        /// <summary>
        /// Get or set inside vertical table cell border style
        /// </summary>
        public WordBorderStyle? InsideVerticalStyle {
            get {
                return BordersOrNull?.InsideVerticalBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var vb = EnsureInsideVertical();
                vb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using hex color codes
        /// </summary>
        public string? InsideVerticalColorHex {
            get {
                return BordersOrNull?.InsideVerticalBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var vb = EnsureInsideVertical();
                vb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set inside vertical table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor InsideVerticalColor {
            get {
                var hex = InsideVerticalColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("InsideVerticalColorHex is null"));
            }
            set { this.InsideVerticalColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the inside vertical cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideVerticalSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.InsideVerticalBorder?.Space?.Value);
            }
                set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var vb = EnsureInsideVertical();
                vb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the inside vertical cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideVerticalSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.InsideVerticalBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var vb = EnsureInsideVertical();
                vb.Size = nativeValue;
            }
        }






        /// <summary>
        /// Get or set start table cell border style
        /// </summary>
        public WordBorderStyle? StartStyle {
            get {
                return BordersOrNull?.StartBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var sb = EnsureStart();
                sb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set start table cell border color using hex color codes
        /// </summary>
        public string? StartColorHex {
            get {
                return BordersOrNull?.StartBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var sb = EnsureStart();
                sb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set start table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor StartColor {
            get {
                var hex = StartColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("StartColorHex is null"));
            }
            set { this.StartColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the start cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? StartSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.StartBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var sb = EnsureStart();
                sb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the start cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? StartSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.StartBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var sb = EnsureStart();
                sb.Size = nativeValue;
            }
        }






        /// <summary>
        /// Get or set end table cell border style
        /// </summary>
        public WordBorderStyle? EndStyle {
            get {
                return BordersOrNull?.EndBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var eb = EnsureEnd();
                eb.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set end table cell border color using hex color codes
        /// </summary>
        public string? EndColorHex {
            get {
                return BordersOrNull?.EndBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var eb = EnsureEnd();
                eb.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set end table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor EndColor {
            get {
                var hex = EndColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("EndColorHex is null"));
            }
            set { this.EndColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the end cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? EndSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.EndBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var eb = EnsureEnd();
                eb.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the end cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? EndSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.EndBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var eb = EnsureEnd();
                eb.Size = nativeValue;
            }
        }













        /// <summary>
        /// Get or set top left to bottom right table cell border style
        /// </summary>
        public WordBorderStyle? TopLeftToBottomRightStyle {
            get {
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var d = EnsureTLBR();
                d.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using hex color codes
        /// </summary>
        public string? TopLeftToBottomRightColorHex {
            get {
                return BordersOrNull?.TopLeftToBottomRightCellBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var d = EnsureTLBR();
                d.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set top left to bottom right table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor TopLeftToBottomRightColor {
            get {
                var hex = TopLeftToBottomRightColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopLeftToBottomRightColorHex is null"));
            }
            set { this.TopLeftToBottomRightColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopLeftToBottomRightSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.TopLeftToBottomRightCellBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var d = EnsureTLBR();
                d.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopLeftToBottomRightSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.TopLeftToBottomRightCellBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var d = EnsureTLBR();
                d.Size = nativeValue;
            }
        }




        /// <summary>
        /// Get or set top right to bottom left table cell border style
        /// </summary>
        public WordBorderStyle? TopRightToBottomLeftStyle {
            get {
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Val?.Value.ToOfficeEnum();
            }
            set {
                var d = EnsureTRBL();
                d.Val = value.HasValue ? value.Value.ToOpenXml() : null;
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using hex color codes
        /// </summary>
        public string? TopRightToBottomLeftColorHex {
            get {
                return BordersOrNull?.TopRightToBottomLeftCellBorder?.Color?.Value is string color ? Helpers.NormalizeOpenXmlColor(color) : null;
            }
            set {
                var d = EnsureTRBL();
                d.Color = value == null ? null : new StringValue(Helpers.NormalizeOpenXmlColor(value));
            }
        }

        /// <summary>
        /// Get or set top right to bottom left table cell border color using named colors
        /// </summary>
        public OfficeIMO.OfficeColor TopRightToBottomLeftColor {
            get {
                var hex = TopRightToBottomLeftColorHex;
                return Helpers.ParseColor(hex ?? throw new InvalidOperationException("TopRightToBottomLeftColorHex is null"));
            }
            set { this.TopRightToBottomLeftColorHex = value.ToRgbHex(); }
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopRightToBottomLeftSpace {
            get {
                return WordBorderLengths.FromSpace(BordersOrNull?.TopRightToBottomLeftCellBorder?.Space?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSpace(value);
                var d = EnsureTRBL();
                d.Space = nativeValue;
            }
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopRightToBottomLeftSize {
            get {
                return WordBorderLengths.FromSize(BordersOrNull?.TopRightToBottomLeftCellBorder?.Size?.Value);
            }
            set {
                uint? nativeValue = WordBorderLengths.ToSize(value);
                var d = EnsureTRBL();
                d.Size = nativeValue;
            }
        }

    }
}
