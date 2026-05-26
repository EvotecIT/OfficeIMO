using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable text style for the full text block of a Visio shape.
    /// </summary>
    public sealed class VisioTextStyle {
        /// <summary>Font family name, such as Aptos or Calibri.</summary>
        public string? FontFamily { get; set; }

        /// <summary>Text color.</summary>
        public Color? Color { get; set; }

        /// <summary>Font size in points.</summary>
        public double? Size { get; set; }

        /// <summary>Whether text is bold.</summary>
        public bool? Bold { get; set; }

        /// <summary>Whether text is italic.</summary>
        public bool? Italic { get; set; }

        /// <summary>Whether text is underlined.</summary>
        public bool? Underline { get; set; }

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

        /// <summary>Text block background color.</summary>
        public Color? BackgroundColor { get; set; }

        /// <summary>Text block background transparency, using Visio's 0-100 scale.</summary>
        public double? BackgroundTransparency { get; set; }

        internal int? FontFaceId { get; set; }

        /// <summary>Creates a detached copy of this text style.</summary>
        public VisioTextStyle Clone() {
            return new VisioTextStyle {
                FontFamily = FontFamily,
                Color = Color,
                Size = Size,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
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
                FontFaceId = FontFaceId
            };
        }
    }
}
