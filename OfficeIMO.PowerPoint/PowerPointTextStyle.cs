using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Defines text formatting that can be applied to runs or paragraphs.
    /// </summary>
    public readonly struct PowerPointTextStyle {
        /// <summary>
        /// Creates a new text style instance.
        /// </summary>
        public PowerPointTextStyle(global::OfficeIMO.OfficeLength? fontSize = null, string? fontFamily = null, global::OfficeIMO.OfficeColor? color = null,
            bool? bold = null, bool? italic = null, bool? underline = null, global::OfficeIMO.OfficeColor? backgroundColor = null,
            bool? strikethrough = null)
            : this(fontSize, fontFamily, color, bold, italic, underline, backgroundColor, strikethrough, null) { }

        private PowerPointTextStyle(global::OfficeIMO.OfficeLength? fontSize, string? fontFamily, global::OfficeIMO.OfficeColor? color,
            bool? bold, bool? italic, bool? underline, global::OfficeIMO.OfficeColor? backgroundColor, bool? strikethrough,
            PowerPointUnderlineStyle? underlineStyle) {
            FontSize = fontSize;
            FontFamily = fontFamily;
            Color = color;
            Bold = bold;
            Italic = italic;
            Underline = underlineStyle.HasValue ? underlineStyle.Value != PowerPointUnderlineStyle.None : underline;
            UnderlineStyle = underlineStyle;
            BackgroundColor = backgroundColor;
            Strikethrough = strikethrough;
        }

        /// <summary>
        /// Physical font size, stored in DrawingML hundredths of a point.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize { get; }

        /// <summary>
        /// Font family (Latin).
        /// </summary>
        public string? FontFamily { get; }

        /// <summary>
        /// Text color; alpha below 255 is written as native DrawingML opacity.
        /// </summary>
        public global::OfficeIMO.OfficeColor? Color { get; }

        /// <summary>
        /// Bold formatting.
        /// </summary>
        public bool? Bold { get; }

        /// <summary>
        /// Italic formatting.
        /// </summary>
        public bool? Italic { get; }

        /// <summary>
        /// Underline formatting.
        /// </summary>
        public bool? Underline { get; }

        /// <summary>Precise native underline variant; null leaves the variant unspecified.</summary>
        public PowerPointUnderlineStyle? UnderlineStyle { get; }

        internal bool HasFormatting => FontSize.HasValue || FontFamily != null || Color.HasValue || Bold.HasValue
            || Italic.HasValue || Underline.HasValue || BackgroundColor.HasValue || Strikethrough.HasValue;

        /// <summary>
        /// Strikethrough formatting.
        /// </summary>
        public bool? Strikethrough { get; }

        /// <summary>
        /// Run background (DrawingML highlight) color; alpha below 255 is written as native opacity.
        /// </summary>
        public global::OfficeIMO.OfficeColor? BackgroundColor { get; }

        /// <summary>
        /// A preset for typical slide titles.
        /// </summary>
        public static PowerPointTextStyle Title => new(fontSize: global::OfficeIMO.OfficeLength.Points(32), bold: true);

        /// <summary>
        /// A preset for subtitles.
        /// </summary>
        public static PowerPointTextStyle Subtitle => new(fontSize: global::OfficeIMO.OfficeLength.Points(24));

        /// <summary>
        /// A preset for body text.
        /// </summary>
        public static PowerPointTextStyle Body => new(fontSize: global::OfficeIMO.OfficeLength.Points(18));

        /// <summary>
        /// A preset for captions or footnotes.
        /// </summary>
        public static PowerPointTextStyle Caption => new(fontSize: global::OfficeIMO.OfficeLength.Points(12));

        /// <summary>
        /// A preset that enables bold emphasis.
        /// </summary>
        public static PowerPointTextStyle Emphasis => new(bold: true);

        /// <summary>
        /// Applies the style to a text run.
        /// </summary>
        public void Apply(PowerPointTextRun run) {
            if (run == null) {
                throw new ArgumentNullException(nameof(run));
            }

            if (FontSize != null) {
                run.FontSize = FontSize.Value;
            }
            if (FontFamily != null) {
                run.FontFamily = FontFamily;
            }
            if (Color.HasValue) {
                run.Color = Color;
            }
            if (Bold != null) {
                run.Bold = Bold.Value;
            }
            if (Italic != null) {
                run.Italic = Italic.Value;
            }
            if (UnderlineStyle.HasValue) run.UnderlineStyle = UnderlineStyle;
            else if (Underline != null) run.Underline = Underline.Value;
            if (Strikethrough != null) {
                run.Strikethrough = Strikethrough.Value;
            }
            if (BackgroundColor.HasValue) {
                run.BackgroundColor = BackgroundColor;
            }
        }

        /// <summary>
        /// Applies the style to a paragraph's default run.
        /// </summary>
        public void Apply(PowerPointParagraph paragraph) {
            if (paragraph == null) {
                throw new ArgumentNullException(nameof(paragraph));
            }

            if (FontSize != null) {
                paragraph.SetFontSize(FontSize.Value);
            }
            if (FontFamily != null) {
                paragraph.SetFontFamily(FontFamily);
            }
            if (Color.HasValue) {
                paragraph.SetColor(Color.Value);
            }
            if (Bold != null) {
                paragraph.SetBold(Bold.Value);
            }
            if (Italic != null) {
                paragraph.SetItalic(Italic.Value);
            }
            if (UnderlineStyle.HasValue) {
                PowerPointUnderlineStyle underlineStyle = UnderlineStyle.Value;
                paragraph.FormatRuns(run => run.UnderlineStyle = underlineStyle);
            } else if (Underline != null) paragraph.SetUnderline(Underline.Value);
            if (Strikethrough != null) {
                paragraph.SetStrikethrough(Strikethrough.Value);
            }
            if (BackgroundColor.HasValue) {
                paragraph.SetBackgroundColor(BackgroundColor.Value);
            }
        }

        /// <summary>
        /// Returns a copy with a new font size.
        /// </summary>
        public PowerPointTextStyle WithFontSize(global::OfficeIMO.OfficeLength? fontSize) {
            return new PowerPointTextStyle(fontSize, FontFamily, Color, Bold, Italic, Underline, BackgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new font family.
        /// </summary>
        public PowerPointTextStyle WithFontFamily(string? fontFamily) {
            return new PowerPointTextStyle(FontSize, fontFamily, Color, Bold, Italic, Underline, BackgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new color.
        /// </summary>
        public PowerPointTextStyle WithColor(global::OfficeIMO.OfficeColor? color) {
            return new PowerPointTextStyle(FontSize, FontFamily, color, Bold, Italic, Underline, BackgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with bold formatting updated.
        /// </summary>
        public PowerPointTextStyle WithBold(bool? bold) {
            return new PowerPointTextStyle(FontSize, FontFamily, Color, bold, Italic, Underline, BackgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with italic formatting updated.
        /// </summary>
        public PowerPointTextStyle WithItalic(bool? italic) {
            return new PowerPointTextStyle(FontSize, FontFamily, Color, Bold, italic, Underline, BackgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with underline formatting updated.
        /// </summary>
        public PowerPointTextStyle WithUnderline(bool? underline) {
            return new PowerPointTextStyle(FontSize, FontFamily, Color, Bold, Italic, underline, BackgroundColor, Strikethrough);
        }

        /// <summary>
        /// Returns a copy with strikethrough formatting updated.
        /// </summary>
        public PowerPointTextStyle WithStrikethrough(bool? strikethrough) {
            return new PowerPointTextStyle(FontSize, FontFamily, Color, Bold, Italic, Underline, BackgroundColor, strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new run background (highlight) color.
        /// </summary>
        public PowerPointTextStyle WithBackgroundColor(global::OfficeIMO.OfficeColor? backgroundColor) {
            return new PowerPointTextStyle(FontSize, FontFamily, Color, Bold, Italic, Underline, backgroundColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>Returns a copy with a native underline variant; null clears the underline override.</summary>
        public PowerPointTextStyle WithUnderlineStyle(PowerPointUnderlineStyle? underlineStyle) {
            if (underlineStyle.HasValue && !Enum.IsDefined(typeof(PowerPointUnderlineStyle), underlineStyle.Value))
                throw new ArgumentOutOfRangeException(nameof(underlineStyle));
            return new PowerPointTextStyle(FontSize, FontFamily, Color, Bold, Italic, null, BackgroundColor, Strikethrough, underlineStyle);
        }

        /// <summary>Applies requested formatting to the paragraph defaults inherited by new placeholder text.</summary>
        internal void ApplyDefaults(A.DefaultRunProperties properties) {
            if (FontSize.HasValue) properties.FontSize = PowerPointTextDefaults.ToDrawingFontSize(FontSize.Value.ToPoints(), nameof(FontSize));
            if (FontFamily != null) PowerPointTextPropertyFormatting.SetFontFamily(properties, FontFamily);
            if (Color.HasValue) PowerPointTextPropertyFormatting.SetColor(properties, Color);
            if (Bold.HasValue) properties.Bold = Bold.Value;
            if (Italic.HasValue) properties.Italic = Italic.Value;
            if (UnderlineStyle.HasValue) properties.Underline = UnderlineStyle.Value.ToOpenXml();
            else if (Underline.HasValue) properties.Underline = Underline.Value ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None;
            if (Strikethrough.HasValue) properties.Strike = Strikethrough.Value ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;
            if (BackgroundColor.HasValue) PowerPointTextPropertyFormatting.SetBackgroundColor(properties, BackgroundColor);
        }
    }
}
