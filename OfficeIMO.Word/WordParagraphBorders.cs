using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word {
    /// <summary>Specifies which side of the paragraph border is affected.</summary>
    public enum WordParagraphBorderType {
        /// <summary>Left border.</summary>
        Left,
        /// <summary>Right border.</summary>
        Right,
        /// <summary>Top border.</summary>
        Top,
        /// <summary>Bottom border.</summary>
        Bottom
    }

    /// <summary>Provides access to paragraph border properties.</summary>
    public class WordParagraphBorders {
        private const int W = OfficeXmlNamespaces.WordprocessingML;

        // CT_PBdr also allows between and bar edges after the four exposed sides.
        private static readonly OfficeXmlChildOrder EdgeOrder = new(
            [W, W, W, W, W, W], ["top", "left", "bottom", "right", "between", "bar"]);
        private static readonly string[] Sides = ["top", "left", "bottom", "right"];

        private readonly Func<OfficeXmlNode> _readProperties;
        private readonly Func<bool, OfficeXmlNode> _editProperties;

        /// <summary>Edits borders in the paragraph properties supplied by a paragraph or a style.</summary>
        /// <param name="readProperties">Returns the existing paragraph properties, if any.</param>
        /// <param name="editProperties">Returns editable paragraph properties. The argument is true when creating or replacing the border container.</param>
        internal WordParagraphBorders(Func<OfficeXmlNode> readProperties, Func<bool, OfficeXmlNode> editProperties) {
            _readProperties = readProperties;
            _editProperties = editProperties;
        }

        // Direct properties leave extension lists and opaque markup-compatibility wrappers untouched.
        private OfficeXmlNode ReadBorders() => _readProperties().Child(W, "pBdr");

        private OfficeXmlNode ReadEdge(string side) => ReadBorders().Child(W, side);

        private OfficeXmlNode EditEdge(string side) {
            OfficeXmlNode borders = _editProperties(ReadBorders().IsNone).GetOrAddChild(W, "pBdr", WPPrView.Order);
            return borders.GetOrAddChild(W, side, EdgeOrder);
        }

        private void SetEdgeAttribute(string side, string attribute, string? value) =>
            EditEdge(side).SetAttribute(W, attribute, value);

        private static string? NormalizeColor(string? color) => color?.Replace("#", string.Empty).ToUpperInvariant();

        private static string? FormatOnOff(bool? value) => value switch {
            true => "true",
            false => "false",
            null => null
        };

        private TEnum? ReadEdgeToken<TEnum>(string side, string attribute, TokenParser<TEnum> parse)
            where TEnum : struct, Enum {
            string? token = ReadEdge(side).Attribute(W, attribute);
            return token == null ? null : WordXmlEdits.Token(parse(token, out TEnum value), value, token);
        }

        private delegate bool TokenParser<TEnum>(string token, out TEnum value) where TEnum : struct, Enum;

        /// <summary>Gets or sets the left border width as a physical length.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSize {
            get => WordBorderLengths.FromSize(ReadEdge("left").WUInt32("sz"));
            set => SetEdgeAttribute("left", "sz", OfficeXmlViewValues.Format(WordBorderLengths.ToSize(value)));
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? LeftColorValue {
            get => NormalizeColor(ReadEdge("left").Attribute(W, "color"));
            set => SetEdgeAttribute("left", "color", NormalizeColor(value));
        }

        /// <summary>Gets or sets the left border color.</summary>
        public OfficeColor? LeftColor {
            get => Helpers.ParseNativeColor(LeftColorValue);
            set => LeftColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the left border theme color.</summary>
        public WordThemeColor? LeftThemeColor {
            get => ReadEdgeToken<WordThemeColor>("left", "themeColor", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("left", "themeColor", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the left border style.</summary>
        public WordBorderStyle? LeftStyle {
            get => ReadEdgeToken<WordBorderStyle>("left", "val", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("left", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the left border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? LeftSpace {
            get => WordBorderLengths.FromSpace(ReadEdge("left").WUInt32("space"));
            set => SetEdgeAttribute("left", "space", OfficeXmlViewValues.Format(WordBorderLengths.ToSpace(value)));
        }

        /// <summary>Gets or sets a value indicating whether the left border has a shadow.</summary>
        public bool? LeftShadow {
            get => ReadEdge("left").OnOffAttr(W, "shadow");
            set => SetEdgeAttribute("left", "shadow", FormatOnOff(value));
        }

        /// <summary>Gets or sets a value indicating whether the left border is part of a frame.</summary>
        public bool? LeftFrame {
            get => ReadEdge("left").OnOffAttr(W, "frame");
            set => SetEdgeAttribute("left", "frame", FormatOnOff(value));
        }

        /// <summary>Gets or sets the right border width as a physical length.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSize {
            get => WordBorderLengths.FromSize(ReadEdge("right").WUInt32("sz"));
            set => SetEdgeAttribute("right", "sz", OfficeXmlViewValues.Format(WordBorderLengths.ToSize(value)));
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? RightColorValue {
            get => NormalizeColor(ReadEdge("right").Attribute(W, "color"));
            set => SetEdgeAttribute("right", "color", NormalizeColor(value));
        }

        /// <summary>Gets or sets the right border color.</summary>
        public OfficeColor? RightColor {
            get => Helpers.ParseNativeColor(RightColorValue);
            set => RightColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the right border theme color.</summary>
        public WordThemeColor? RightThemeColor {
            get => ReadEdgeToken<WordThemeColor>("right", "themeColor", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("right", "themeColor", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the right border style.</summary>
        public WordBorderStyle? RightStyle {
            get => ReadEdgeToken<WordBorderStyle>("right", "val", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("right", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the right border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? RightSpace {
            get => WordBorderLengths.FromSpace(ReadEdge("right").WUInt32("space"));
            set => SetEdgeAttribute("right", "space", OfficeXmlViewValues.Format(WordBorderLengths.ToSpace(value)));
        }

        /// <summary>Gets or sets a value indicating whether the right border has a shadow.</summary>
        public bool? RightShadow {
            get => ReadEdge("right").OnOffAttr(W, "shadow");
            set => SetEdgeAttribute("right", "shadow", FormatOnOff(value));
        }

        /// <summary>Gets or sets a value indicating whether the right border is part of a frame.</summary>
        public bool? RightFrame {
            get => ReadEdge("right").OnOffAttr(W, "frame");
            set => SetEdgeAttribute("right", "frame", FormatOnOff(value));
        }

        /// <summary>Gets or sets the top border width as a physical length.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSize {
            get => WordBorderLengths.FromSize(ReadEdge("top").WUInt32("sz"));
            set => SetEdgeAttribute("top", "sz", OfficeXmlViewValues.Format(WordBorderLengths.ToSize(value)));
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? TopColorValue {
            get => NormalizeColor(ReadEdge("top").Attribute(W, "color"));
            set => SetEdgeAttribute("top", "color", NormalizeColor(value));
        }

        /// <summary>Gets or sets the top border color.</summary>
        public OfficeColor? TopColor {
            get => Helpers.ParseNativeColor(TopColorValue);
            set => TopColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the top border theme color.</summary>
        public WordThemeColor? TopThemeColor {
            get => ReadEdgeToken<WordThemeColor>("top", "themeColor", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("top", "themeColor", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the top border style.</summary>
        public WordBorderStyle? TopStyle {
            get => ReadEdgeToken<WordBorderStyle>("top", "val", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("top", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the top border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? TopSpace {
            get => WordBorderLengths.FromSpace(ReadEdge("top").WUInt32("space"));
            set => SetEdgeAttribute("top", "space", OfficeXmlViewValues.Format(WordBorderLengths.ToSpace(value)));
        }

        /// <summary>Gets or sets a value indicating whether the top border has a shadow.</summary>
        public bool? TopShadow {
            get => ReadEdge("top").OnOffAttr(W, "shadow");
            set => SetEdgeAttribute("top", "shadow", FormatOnOff(value));
        }

        /// <summary>Gets or sets a value indicating whether the top border is part of a frame.</summary>
        public bool? TopFrame {
            get => ReadEdge("top").OnOffAttr(W, "frame");
            set => SetEdgeAttribute("top", "frame", FormatOnOff(value));
        }

        /// <summary>Gets or sets the bottom border width as a physical length.</summary>
        /// <remarks>Nonnegative lengths round to the nearest eighth of a point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSize {
            get => WordBorderLengths.FromSize(ReadEdge("bottom").WUInt32("sz"));
            set => SetEdgeAttribute("bottom", "sz", OfficeXmlViewValues.Format(WordBorderLengths.ToSize(value)));
        }

        /// <summary>Native border colour value (hexadecimal or <c>auto</c>).</summary>
        internal string? BottomColorValue {
            get => NormalizeColor(ReadEdge("bottom").Attribute(W, "color"));
            set => SetEdgeAttribute("bottom", "color", NormalizeColor(value));
        }

        /// <summary>Gets or sets the bottom border color.</summary>
        public OfficeColor? BottomColor {
            get => Helpers.ParseNativeColor(BottomColorValue);
            set => BottomColorValue = value?.ToRgbHex();
        }

        /// <summary>Gets or sets the bottom border theme color.</summary>
        public WordThemeColor? BottomThemeColor {
            get => ReadEdgeToken<WordThemeColor>("bottom", "themeColor", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("bottom", "themeColor", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the bottom border style.</summary>
        public WordBorderStyle? BottomStyle {
            get => ReadEdgeToken<WordBorderStyle>("bottom", "val", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetEdgeAttribute("bottom", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the bottom border spacing.</summary>
        /// <remarks>Nonnegative lengths round to the nearest whole point. Midpoint ties round away from zero. Null clears the attribute.</remarks>
        public OfficeLength? BottomSpace {
            get => WordBorderLengths.FromSpace(ReadEdge("bottom").WUInt32("space"));
            set => SetEdgeAttribute("bottom", "space", OfficeXmlViewValues.Format(WordBorderLengths.ToSpace(value)));
        }

        /// <summary>Gets or sets a value indicating whether the bottom border has a shadow.</summary>
        public bool? BottomShadow {
            get => ReadEdge("bottom").OnOffAttr(W, "shadow");
            set => SetEdgeAttribute("bottom", "shadow", FormatOnOff(value));
        }

        /// <summary>Gets or sets a value indicating whether the bottom border is part of a frame.</summary>
        public bool? BottomFrame {
            get => ReadEdge("bottom").OnOffAttr(W, "frame");
            set => SetEdgeAttribute("bottom", "frame", FormatOnOff(value));
        }

        private void SetBorder(WordBorder border) {
            switch (border) {
                case WordBorder.None:
                    if (!ReadBorders().IsNone) _editProperties(false).Child(W, "pBdr").Remove();
                    return;
                case WordBorder.Custom:
                case WordBorder.Box:
                case WordBorder.Shadow:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(border));
            }

            OfficeXmlNode properties = _editProperties(true);
            properties.RemoveChildren(W, "pBdr");
            OfficeXmlNode borders = properties.AddOrdered(W, "pBdr", WPPrView.Order);
            if (border == WordBorder.Custom) return;

            foreach (string side in Sides) {
                OfficeXmlNode edge = borders.Append(borders.CreateElement(W, side));
                edge.SetAttribute(W, "val", "single");
                edge.SetAttribute(W, "color", "auto");
                edge.SetAttribute(W, "sz", "4");
                edge.SetAttribute(W, "space", "24");
                if (border == WordBorder.Shadow) edge.SetAttribute(W, "shadow", "true");
            }
        }

        /// <summary>Gets or sets the current border preset type.</summary>
        public WordBorder Type {
            get {
                OfficeXmlNode borders = ReadBorders();
                if (borders.IsNone) return WordBorder.None;

                // Presets are classified by the four sides' color, size, spacing and shadow.
                bool? shadow = borders.Child(W, "top").OnOffAttr(W, "shadow");
                if (shadow == false) return WordBorder.Custom;

                foreach (string side in Sides) {
                    OfficeXmlNode edge = borders.Child(W, side);
                    if (edge.Attribute(W, "color") != "auto" || edge.WUInt32("sz") != 4 || edge.WUInt32("space") != 24
                        || edge.OnOffAttr(W, "shadow") != shadow) return WordBorder.Custom;
                }

                return shadow == true ? WordBorder.Shadow : WordBorder.Box;
            }
            set => SetBorder(value);
        }

        /// <summary>Sets the style, color, width and shadow of one paragraph border.</summary>
        /// <param name="type">Border side.</param>
        /// <param name="style">Border style.</param>
        /// <param name="color">Border color.</param>
        /// <param name="size">Border width, rounded to the nearest eighth of a point.</param>
        /// <param name="shadow">Whether the border has a shadow.</param>
        public void SetBorder(WordParagraphBorderType type, WordBorderStyle style, Color color, OfficeLength size, bool shadow) {
            uint? nativeSize = WordBorderLengths.ToSize(size);
            string? side = type switch {
                WordParagraphBorderType.Left => "left",
                WordParagraphBorderType.Right => "right",
                WordParagraphBorderType.Top => "top",
                WordParagraphBorderType.Bottom => "bottom",
                _ => null
            };
            if (side == null) return;

            OfficeXmlNode edge = EditEdge(side);
            edge.SetAttribute(W, "val", style.ToOpenXmlToken());
            edge.SetAttribute(W, "color", color.ToRgbHex());
            edge.SetAttribute(W, "sz", OfficeXmlViewValues.Format(nativeSize));
            edge.SetAttribute(W, "shadow", FormatOnOff(shadow));
        }
    }
}
