using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides a rich API for configuring the borders of a
    /// <see cref="WordTableCell"/>.  Every edge of the cell—including the
    /// diagonals—has properties for style, color, spacing and size so that
    /// callers can individually tailor each side of a cell.
    /// </summary>
    /// <remarks>When creating an edge, assign its style before its colour, spacing or size; Word requires a style on each border descriptor.</remarks>
    public class WordTableCellBorder {
        private const int W = OfficeXmlNamespaces.WordprocessingML;

        internal enum Edge { Left, Right, Top, Bottom, InsideHorizontal, InsideVertical, Start, End, TopLeftToBottomRight, TopRightToBottomLeft }

        // Element names in the order of Edge.
        private static readonly string[] EdgeNames = { "left", "right", "top", "bottom", "insideH", "insideV", "start", "end", "tl2br", "tr2bl" };

        // The schema lists w:start before w:left, but cells have always been written with w:left first (and w:right
        // before w:end) when both are present; keep that order so documents stay byte-for-byte the same.
        private static readonly OfficeXmlChildOrder BorderOrder = new(
            new[] { W, W, W, W, W, W, W, W, W, W },
            new[] { "top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl" });

        private readonly WordTableCell _wordTableCell;
        private readonly WordDocument _document;

        internal WordTableCellBorder(WordDocument wordDocument, WordTableCell wordTableCell) {
            _document = wordDocument;
            _wordTableCell = wordTableCell;
            // Reading never creates cell properties; only the setters add them.
        }

        // The edge is a direct child of the cell's own w:tcBorders, never one inside extension markup.
        private OfficeXmlNode Border(Edge edge) =>
            OfficeXmlNode.FromSdk(_wordTableCell._tableCellProperties).Child(W, "tcBorders").Child(W, EdgeNames[(int)edge]);

        private OfficeXmlNode EnsureBorder(Edge edge) {
            if (_wordTableCell._tableCellProperties == null) {
                _wordTableCell.AddTableCellProperties();
            }
            OfficeXmlNode properties = OfficeXmlNode.FromSdk(_wordTableCell._tableCellProperties);
            OfficeXmlNode borders = properties.Child(W, "tcBorders");
            if (borders.IsNone) {
                borders = properties.AddOrdered(W, "tcBorders", WTcPrView.Order);
                _wordTableCell.NormalizeTableCellPropertiesOrder();
            }
            return borders.GetOrAddChild(W, EdgeNames[(int)edge], BorderOrder);
        }

        internal WordBorderStyle? GetStyle(Edge edge) =>
            Border(edge).Attribute(W, "val") is string token && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordBorderStyle style) ? style : null;

        internal void SetStyle(Edge edge, WordBorderStyle? value) {
            string? token = value?.ToOpenXmlToken();
            EnsureBorder(edge).SetAttribute(W, "val", token);
        }

        private string? GetColorValue(Edge edge) {
            OfficeXmlNode border = Border(edge);
            if (border.IsNone) return null;
            // One pass over the attributes: asking for each one in turn is slower when most of them are absent.
            string? color = null, themeColor = null, themeTint = null, themeShade = null;
            foreach (OfficeXmlAttribute attribute in border.Attributes()) {
                if (!OfficeXmlNamespaces.TryGetKnownId(attribute.NamespaceUri, out int namespaceId) || namespaceId != W) continue;
                switch (attribute.LocalName) {
                    case "color": color = attribute.Value; break;
                    case "themeColor": themeColor = attribute.Value; break;
                    case "themeTint": themeTint = attribute.Value; break;
                    case "themeShade": themeShade = attribute.Value; break;
                }
            }
            string? resolved = (themeColor == null ? null : WordRunFormatting.ResolveThemeColorHex(themeColor, themeTint, themeShade, _document)) ?? color;
            return resolved == null ? null : Helpers.NormalizeOpenXmlColor(resolved);
        }

        /// <summary>A literal color replaces theme color references, as Word's direct color selection does.</summary>
        internal void SetColorValue(Edge edge, string? value) {
            string? color = value == null ? null : Helpers.NormalizeOpenXmlColor(value);
            OfficeXmlNode border = EnsureBorder(edge);
            border.SetAttribute(W, "color", color);
            border.SetAttribute(W, "themeColor", null);
            border.SetAttribute(W, "themeTint", null);
            border.SetAttribute(W, "themeShade", null);
        }

        private OfficeLength? GetSpace(Edge edge) => WordBorderLengths.FromSpace(OfficeXmlViewValues.ParseUInt32(Border(edge).Attribute(W, "space")));

        internal void SetSpace(Edge edge, OfficeLength? value) {
            uint? nativeValue = WordBorderLengths.ToSpace(value);
            EnsureBorder(edge).SetAttribute(W, "space", OfficeXmlViewValues.Format(nativeValue));
        }

        private OfficeLength? GetSize(Edge edge) => WordBorderLengths.FromSize(OfficeXmlViewValues.ParseUInt32(Border(edge).Attribute(W, "sz")));

        internal void SetSize(Edge edge, OfficeLength? value) {
            uint? nativeValue = WordBorderLengths.ToSize(value);
            EnsureBorder(edge).SetAttribute(W, "sz", OfficeXmlViewValues.Format(nativeValue));
        }

        /// <summary>Gets or sets the left cell border style.</summary>
        public WordBorderStyle? LeftStyle {
            get => GetStyle(Edge.Left);
            set => SetStyle(Edge.Left, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? LeftColorValue {
            get => GetColorValue(Edge.Left);
            set => SetColorValue(Edge.Left, value);
        }

        /// <summary>Gets or sets the left cell border color.</summary>
        public OfficeIMO.OfficeColor? LeftColor {
            get => Helpers.ParseNativeColor(LeftColorValue);
            set => LeftColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the left cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSpace {
            get => GetSpace(Edge.Left);
            set => SetSpace(Edge.Left, value);
        }

        /// <summary>Gets or sets the left cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSize {
            get => GetSize(Edge.Left);
            set => SetSize(Edge.Left, value);
        }

        /// <summary>Gets or sets the right cell border style.</summary>
        public WordBorderStyle? RightStyle {
            get => GetStyle(Edge.Right);
            set => SetStyle(Edge.Right, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? RightColorValue {
            get => GetColorValue(Edge.Right);
            set => SetColorValue(Edge.Right, value);
        }

        /// <summary>Gets or sets the right cell border color.</summary>
        public OfficeIMO.OfficeColor? RightColor {
            get => Helpers.ParseNativeColor(RightColorValue);
            set => RightColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the right cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSpace {
            get => GetSpace(Edge.Right);
            set => SetSpace(Edge.Right, value);
        }

        /// <summary>Gets or sets the right cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSize {
            get => GetSize(Edge.Right);
            set => SetSize(Edge.Right, value);
        }

        /// <summary>Gets or sets the top cell border style.</summary>
        public WordBorderStyle? TopStyle {
            get => GetStyle(Edge.Top);
            set => SetStyle(Edge.Top, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? TopColorValue {
            get => GetColorValue(Edge.Top);
            set => SetColorValue(Edge.Top, value);
        }

        /// <summary>Gets or sets the top cell border color.</summary>
        public OfficeIMO.OfficeColor? TopColor {
            get => Helpers.ParseNativeColor(TopColorValue);
            set => TopColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the top cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSpace {
            get => GetSpace(Edge.Top);
            set => SetSpace(Edge.Top, value);
        }

        /// <summary>Gets or sets the top cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSize {
            get => GetSize(Edge.Top);
            set => SetSize(Edge.Top, value);
        }

        /// <summary>Gets or sets the bottom cell border style.</summary>
        public WordBorderStyle? BottomStyle {
            get => GetStyle(Edge.Bottom);
            set => SetStyle(Edge.Bottom, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? BottomColorValue {
            get => GetColorValue(Edge.Bottom);
            set => SetColorValue(Edge.Bottom, value);
        }

        /// <summary>Gets or sets the bottom cell border color.</summary>
        public OfficeIMO.OfficeColor? BottomColor {
            get => Helpers.ParseNativeColor(BottomColorValue);
            set => BottomColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the bottom cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSpace {
            get => GetSpace(Edge.Bottom);
            set => SetSpace(Edge.Bottom, value);
        }

        /// <summary>Gets or sets the bottom cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSize {
            get => GetSize(Edge.Bottom);
            set => SetSize(Edge.Bottom, value);
        }

        /// <summary>Gets or sets the inside horizontal cell border style.</summary>
        public WordBorderStyle? InsideHorizontalStyle {
            get => GetStyle(Edge.InsideHorizontal);
            set => SetStyle(Edge.InsideHorizontal, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? InsideHorizontalColorValue {
            get => GetColorValue(Edge.InsideHorizontal);
            set => SetColorValue(Edge.InsideHorizontal, value);
        }

        /// <summary>Gets or sets the inside horizontal cell border color.</summary>
        public OfficeIMO.OfficeColor? InsideHorizontalColor {
            get => Helpers.ParseNativeColor(InsideHorizontalColorValue);
            set => InsideHorizontalColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the inside horizontal cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideHorizontalSpace {
            get => GetSpace(Edge.InsideHorizontal);
            set => SetSpace(Edge.InsideHorizontal, value);
        }

        /// <summary>Gets or sets the inside horizontal cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideHorizontalSize {
            get => GetSize(Edge.InsideHorizontal);
            set => SetSize(Edge.InsideHorizontal, value);
        }

        /// <summary>Gets or sets the inside vertical cell border style.</summary>
        public WordBorderStyle? InsideVerticalStyle {
            get => GetStyle(Edge.InsideVertical);
            set => SetStyle(Edge.InsideVertical, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? InsideVerticalColorValue {
            get => GetColorValue(Edge.InsideVertical);
            set => SetColorValue(Edge.InsideVertical, value);
        }

        /// <summary>Gets or sets the inside vertical cell border color.</summary>
        public OfficeIMO.OfficeColor? InsideVerticalColor {
            get => Helpers.ParseNativeColor(InsideVerticalColorValue);
            set => InsideVerticalColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the inside vertical cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideVerticalSpace {
            get => GetSpace(Edge.InsideVertical);
            set => SetSpace(Edge.InsideVertical, value);
        }

        /// <summary>Gets or sets the inside vertical cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? InsideVerticalSize {
            get => GetSize(Edge.InsideVertical);
            set => SetSize(Edge.InsideVertical, value);
        }

        /// <summary>Gets or sets the start cell border style.</summary>
        public WordBorderStyle? StartStyle {
            get => GetStyle(Edge.Start);
            set => SetStyle(Edge.Start, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? StartColorValue {
            get => GetColorValue(Edge.Start);
            set => SetColorValue(Edge.Start, value);
        }

        /// <summary>Gets or sets the start cell border color.</summary>
        public OfficeIMO.OfficeColor? StartColor {
            get => Helpers.ParseNativeColor(StartColorValue);
            set => StartColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the start cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? StartSpace {
            get => GetSpace(Edge.Start);
            set => SetSpace(Edge.Start, value);
        }

        /// <summary>Gets or sets the start cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? StartSize {
            get => GetSize(Edge.Start);
            set => SetSize(Edge.Start, value);
        }

        /// <summary>Gets or sets the end cell border style.</summary>
        public WordBorderStyle? EndStyle {
            get => GetStyle(Edge.End);
            set => SetStyle(Edge.End, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? EndColorValue {
            get => GetColorValue(Edge.End);
            set => SetColorValue(Edge.End, value);
        }

        /// <summary>Gets or sets the end cell border color.</summary>
        public OfficeIMO.OfficeColor? EndColor {
            get => Helpers.ParseNativeColor(EndColorValue);
            set => EndColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the end cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? EndSpace {
            get => GetSpace(Edge.End);
            set => SetSpace(Edge.End, value);
        }

        /// <summary>Gets or sets the end cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? EndSize {
            get => GetSize(Edge.End);
            set => SetSize(Edge.End, value);
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border style.</summary>
        public WordBorderStyle? TopLeftToBottomRightStyle {
            get => GetStyle(Edge.TopLeftToBottomRight);
            set => SetStyle(Edge.TopLeftToBottomRight, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? TopLeftToBottomRightColorValue {
            get => GetColorValue(Edge.TopLeftToBottomRight);
            set => SetColorValue(Edge.TopLeftToBottomRight, value);
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border color.</summary>
        public OfficeIMO.OfficeColor? TopLeftToBottomRightColor {
            get => Helpers.ParseNativeColor(TopLeftToBottomRightColorValue);
            set => TopLeftToBottomRightColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopLeftToBottomRightSpace {
            get => GetSpace(Edge.TopLeftToBottomRight);
            set => SetSpace(Edge.TopLeftToBottomRight, value);
        }

        /// <summary>Gets or sets the top-left to bottom-right cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopLeftToBottomRightSize {
            get => GetSize(Edge.TopLeftToBottomRight);
            set => SetSize(Edge.TopLeftToBottomRight, value);
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border style.</summary>
        public WordBorderStyle? TopRightToBottomLeftStyle {
            get => GetStyle(Edge.TopRightToBottomLeft);
            set => SetStyle(Edge.TopRightToBottomLeft, value);
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? TopRightToBottomLeftColorValue {
            get => GetColorValue(Edge.TopRightToBottomLeft);
            set => SetColorValue(Edge.TopRightToBottomLeft, value);
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border color.</summary>
        public OfficeIMO.OfficeColor? TopRightToBottomLeftColor {
            get => Helpers.ParseNativeColor(TopRightToBottomLeftColorValue);
            set => TopRightToBottomLeftColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopRightToBottomLeftSpace {
            get => GetSpace(Edge.TopRightToBottomLeft);
            set => SetSpace(Edge.TopRightToBottomLeft, value);
        }

        /// <summary>Gets or sets the top-right to bottom-left cell border width.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point; midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopRightToBottomLeftSize {
            get => GetSize(Edge.TopRightToBottomLeft);
            set => SetSize(Edge.TopRightToBottomLeft, value);
        }
    }
}
