using System;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable text style for the full text block of a Visio shape.
    /// </summary>
    public sealed class VisioTextStyle {
        private OfficeTextDecorationStyle? _underlineStyle;
        private OfficeTextDecorationStyle? _strikethroughStyle;
        private VisioTextCapitalization? _capitalization;
        private OfficeTextBaseline? _baseline;
        private string? _fontFamily;
        /// <summary>Font family name, such as Aptos or Calibri.</summary>
        public string? FontFamily {
            get => _fontFamily;
            set { _fontFamily = value; FontFamilyAssigned = true; }
        }
        internal bool FontFamilyAssigned { get; set; }

        /// <summary>Text color.</summary>
        public Color? Color { get; set; }

        /// <summary>Font size in points.</summary>
        public double? Size { get; set; }

        /// <summary>Whether text is bold.</summary>
        public bool? Bold { get; set; }

        /// <summary>Whether text is italic.</summary>
        public bool? Italic { get; set; }

        /// <summary>Whether text is underlined. This compatibility property maps to a single underline.</summary>
        public bool? Underline {
            get => _underlineStyle.HasValue ? _underlineStyle.Value != OfficeTextDecorationStyle.None : (bool?)null;
            set => _underlineStyle = value.HasValue
                ? value.Value ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None
                : (OfficeTextDecorationStyle?)null;
        }

        /// <summary>Native underline variant. Visio supports none, single, and double.</summary>
        public OfficeTextDecorationStyle? UnderlineStyle {
            get => _underlineStyle;
            set {
                ValidateNativeDecoration(value, nameof(UnderlineStyle));
                _underlineStyle = value;
            }
        }

        /// <summary>Whether text has strikethrough. This compatibility property maps to a single strike line.</summary>
        public bool? Strikethrough {
            get => _strikethroughStyle.HasValue ? _strikethroughStyle.Value != OfficeTextDecorationStyle.None : (bool?)null;
            set => _strikethroughStyle = value.HasValue
                ? value.Value ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None
                : (OfficeTextDecorationStyle?)null;
        }

        /// <summary>Native strikethrough variant. Visio supports none, single, and double.</summary>
        public OfficeTextDecorationStyle? StrikethroughStyle {
            get => _strikethroughStyle;
            set {
                ValidateNativeDecoration(value, nameof(StrikethroughStyle));
                _strikethroughStyle = value;
            }
        }

        /// <summary>Whether Visio's native small-capital style bit is active.</summary>
        public bool? SmallCaps { get; set; }

        /// <summary>Native display-time capitalization mode.</summary>
        public VisioTextCapitalization? Capitalization {
            get => _capitalization;
            set {
                if (value.HasValue && (value.Value < VisioTextCapitalization.Normal || value.Value > VisioTextCapitalization.InitialCaps)) throw new ArgumentOutOfRangeException(nameof(Capitalization));
                _capitalization = value;
            }
        }

        /// <summary>Native baseline placement.</summary>
        public OfficeTextBaseline? Baseline {
            get => _baseline;
            set {
                if (value.HasValue && (value.Value < OfficeTextBaseline.Normal || value.Value > OfficeTextBaseline.Subscript)) throw new ArgumentOutOfRangeException(nameof(Baseline));
                _baseline = value;
            }
        }

        /// <summary>Horizontal text alignment.</summary>
        public VisioTextHorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>Vertical text alignment.</summary>
        public VisioTextVerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>Left text margin in inches.</summary>
        public double? LeftMargin { get; set; }

        /// <summary>Right text margin in inches.</summary>
        public double? RightMargin { get; set; }

        /// <summary>Top text margin in inches.</summary>
        public double? TopMargin { get; set; }

        /// <summary>Bottom text margin in inches.</summary>
        public double? BottomMargin { get; set; }

        /// <summary>Text block pin X relative to the shape origin, in inches.</summary>
        public double? TextPinX { get; set; }

        /// <summary>Text block pin Y relative to the shape origin, in inches.</summary>
        public double? TextPinY { get; set; }

        /// <summary>Text block width in inches.</summary>
        public double? TextWidth { get; set; }

        /// <summary>Text block height in inches.</summary>
        public double? TextHeight { get; set; }

        /// <summary>Text block local pin X relative to the text block origin, in inches.</summary>
        public double? TextLocPinX { get; set; }

        /// <summary>Text block local pin Y relative to the text block origin, in inches.</summary>
        public double? TextLocPinY { get; set; }

        /// <summary>Text block rotation angle in radians.</summary>
        public double? TextAngle { get; set; }

        private Color? _backgroundColor;
        private double? _backgroundTransparency;

        /// <summary>Text block background color. A color with zero alpha explicitly disables the background.</summary>
        public Color? BackgroundColor {
            get => _backgroundColor;
            set { _backgroundColor = value; NativeBackgroundColorCell = null; BackgroundColorAssigned = true; }
        }

        /// <summary>Text block background transparency as a percentage from 0 (opaque) to 100 (transparent).</summary>
        public double? BackgroundTransparency {
            get => _backgroundTransparency;
            set { _backgroundTransparency = value; NativeBackgroundTransparencyCell = null; BackgroundTransparencyAssigned = true; }
        }

        // Explicit assignments, including the same value, replace native formulas and errors.
        // Loader initialization and detached cloning restore source cells after assigning values.
        internal XElement? NativeBackgroundColorCell { get; set; }
        internal XElement? NativeBackgroundTransparencyCell { get; set; }
        internal bool BackgroundColorAssigned { get; set; }
        internal bool BackgroundTransparencyAssigned { get; set; }

        internal int? FontFaceId { get; set; }

        /// <summary>Creates a detached copy of this text style.</summary>
        public VisioTextStyle Clone() {
            return new VisioTextStyle {
                FontFamily = FontFamily,
                Color = Color,
                Size = Size,
                Bold = Bold,
                Italic = Italic,
                UnderlineStyle = UnderlineStyle,
                StrikethroughStyle = StrikethroughStyle,
                SmallCaps = SmallCaps,
                Capitalization = Capitalization,
                Baseline = Baseline,
                HorizontalAlignment = HorizontalAlignment,
                VerticalAlignment = VerticalAlignment,
                LeftMargin = LeftMargin,
                RightMargin = RightMargin,
                TopMargin = TopMargin,
                BottomMargin = BottomMargin,
                TextPinX = TextPinX,
                TextPinY = TextPinY,
                TextWidth = TextWidth,
                TextHeight = TextHeight,
                TextLocPinX = TextLocPinX,
                TextLocPinY = TextLocPinY,
                TextAngle = TextAngle,
                BackgroundColor = BackgroundColor,
                BackgroundTransparency = BackgroundTransparency,
                FontFaceId = FontFaceId,
                FontFamilyAssigned = FontFamilyAssigned,
                NativeBackgroundColorCell = NativeBackgroundColorCell == null ? null : new XElement(NativeBackgroundColorCell),
                NativeBackgroundTransparencyCell = NativeBackgroundTransparencyCell == null ? null : new XElement(NativeBackgroundTransparencyCell),
                BackgroundColorAssigned = BackgroundColorAssigned,
                BackgroundTransparencyAssigned = BackgroundTransparencyAssigned
            };
        }

        /// <summary>Fills unspecified properties from a detached master style.</summary>
        internal void InheritUnsetFrom(VisioTextStyle source, bool inheritCharacter = true, bool inheritParagraph = true) {
            if (inheritCharacter) {
                _fontFamily ??= source.FontFamily;
                Color ??= source.Color;
                Size ??= source.Size;
                Bold ??= source.Bold;
                Italic ??= source.Italic;
                UnderlineStyle ??= source.UnderlineStyle;
                StrikethroughStyle ??= source.StrikethroughStyle;
                SmallCaps ??= source.SmallCaps;
                Capitalization ??= source.Capitalization;
                Baseline ??= source.Baseline;
                FontFaceId ??= source.FontFaceId;
            }
            if (inheritParagraph) HorizontalAlignment ??= source.HorizontalAlignment;
            VerticalAlignment ??= source.VerticalAlignment;
            LeftMargin ??= source.LeftMargin;
            RightMargin ??= source.RightMargin;
            TopMargin ??= source.TopMargin;
            BottomMargin ??= source.BottomMargin;
            TextPinX ??= source.TextPinX;
            TextPinY ??= source.TextPinY;
            TextWidth ??= source.TextWidth;
            TextHeight ??= source.TextHeight;
            TextLocPinX ??= source.TextLocPinX;
            TextLocPinY ??= source.TextLocPinY;
            TextAngle ??= source.TextAngle;
            _backgroundColor ??= source.BackgroundColor;
            _backgroundTransparency ??= source.BackgroundTransparency;
        }

        internal void ScaleTextBlock(double x, double y) {
            TextPinX *= x; TextPinY *= y; TextWidth *= x; TextHeight *= y;
            TextLocPinX *= x; TextLocPinY *= y;
        }

        private static void ValidateNativeDecoration(OfficeTextDecorationStyle? style, string propertyName) {
            if (!style.HasValue) return;
            if (style.Value != OfficeTextDecorationStyle.None &&
                 style.Value != OfficeTextDecorationStyle.Single &&
                 style.Value != OfficeTextDecorationStyle.Double) {
                throw new ArgumentOutOfRangeException(propertyName, style, "Visio supports only none, single, and double text decoration lines.");
            }
        }
    }
}
