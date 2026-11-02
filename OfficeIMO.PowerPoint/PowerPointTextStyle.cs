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
        public PowerPointTextStyle(global::OfficeIMO.OfficeLength? fontSize = null, string? fontName = null, string? color = null,
            bool? bold = null, bool? italic = null, bool? underline = null, string? highlightColor = null,
            bool? strikethrough = null)
            : this(fontSize, fontName, color, bold, italic, underline, highlightColor, strikethrough, null) { }

        private PowerPointTextStyle(global::OfficeIMO.OfficeLength? fontSize, string? fontName, string? color,
            bool? bold, bool? italic, bool? underline, string? highlightColor, bool? strikethrough,
            PowerPointUnderlineStyle? underlineStyle) {
            FontSize = fontSize;
            FontName = fontName;
            Color = color;
            Bold = bold;
            Italic = italic;
            Underline = underlineStyle.HasValue ? underlineStyle.Value != PowerPointUnderlineStyle.None : underline;
            UnderlineStyle = underlineStyle;
            HighlightColor = highlightColor;
            Strikethrough = strikethrough;
        }

        /// <summary>
        /// Physical font size, stored in DrawingML hundredths of a point.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize { get; }

        /// <summary>
        /// Font name (Latin).
        /// </summary>
        public string? FontName { get; }

        /// <summary>
        /// Text color as an RGB or RGBA hexadecimal literal (e.g. "1F4E7980").
        /// </summary>
        public string? Color { get; }

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

        internal bool HasFormatting => FontSize.HasValue || FontName != null || Color != null || Bold.HasValue
            || Italic.HasValue || Underline.HasValue || HighlightColor != null || Strikethrough.HasValue;

        /// <summary>
        /// Strikethrough formatting.
        /// </summary>
        public bool? Strikethrough { get; }

        /// <summary>
        /// Highlight color as an RGB or RGBA hexadecimal literal (e.g. "FFF59D80").
        /// </summary>
        public string? HighlightColor { get; }

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
            if (FontName != null) {
                run.FontName = FontName;
            }
            if (Color != null) {
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
            if (HighlightColor != null) {
                run.HighlightColor = HighlightColor;
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
            if (FontName != null) {
                paragraph.SetFontName(FontName);
            }
            if (Color != null) {
                paragraph.SetColor(Color);
            }
            if (Bold != null) {
                paragraph.SetBold(Bold.Value);
            }
            if (Italic != null) {
                paragraph.SetItalic(Italic.Value);
            }
            if (UnderlineStyle.HasValue) {
                paragraph.SetUnderline(UnderlineStyle.Value != PowerPointUnderlineStyle.None);
                paragraph.Runs[paragraph.Runs.Count - 1].UnderlineStyle = UnderlineStyle;
            } else if (Underline != null) paragraph.SetUnderline(Underline.Value);
            if (Strikethrough != null) {
                paragraph.SetStrikethrough(Strikethrough.Value);
            }
            if (HighlightColor != null) {
                paragraph.SetHighlightColor(HighlightColor);
            }
        }

        /// <summary>
        /// Returns a copy with a new font size.
        /// </summary>
        public PowerPointTextStyle WithFontSize(global::OfficeIMO.OfficeLength? fontSize) {
            return new PowerPointTextStyle(fontSize, FontName, Color, Bold, Italic, Underline, HighlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new font name.
        /// </summary>
        public PowerPointTextStyle WithFontName(string? fontName) {
            return new PowerPointTextStyle(FontSize, fontName, Color, Bold, Italic, Underline, HighlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new color.
        /// </summary>
        public PowerPointTextStyle WithColor(string? color) {
            return new PowerPointTextStyle(FontSize, FontName, color, Bold, Italic, Underline, HighlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with bold formatting updated.
        /// </summary>
        public PowerPointTextStyle WithBold(bool? bold) {
            return new PowerPointTextStyle(FontSize, FontName, Color, bold, Italic, Underline, HighlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with italic formatting updated.
        /// </summary>
        public PowerPointTextStyle WithItalic(bool? italic) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, italic, Underline, HighlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with underline formatting updated.
        /// </summary>
        public PowerPointTextStyle WithUnderline(bool? underline) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, Italic, underline, HighlightColor, Strikethrough);
        }

        /// <summary>
        /// Returns a copy with strikethrough formatting updated.
        /// </summary>
        public PowerPointTextStyle WithStrikethrough(bool? strikethrough) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, Italic, Underline, HighlightColor, strikethrough, UnderlineStyle);
        }

        /// <summary>
        /// Returns a copy with a new highlight color.
        /// </summary>
        public PowerPointTextStyle WithHighlightColor(string? highlightColor) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, Italic, Underline, highlightColor, Strikethrough, UnderlineStyle);
        }

        /// <summary>Returns a copy with a native underline variant; null clears the underline override.</summary>
        public PowerPointTextStyle WithUnderlineStyle(PowerPointUnderlineStyle? underlineStyle) {
            if (underlineStyle.HasValue && !Enum.IsDefined(typeof(PowerPointUnderlineStyle), underlineStyle.Value))
                throw new ArgumentOutOfRangeException(nameof(underlineStyle));
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, Italic, null, HighlightColor, Strikethrough, underlineStyle);
        }

        /// <summary>Applies requested formatting to the paragraph defaults inherited by new placeholder text.</summary>
        internal void ApplyDefaults(A.DefaultRunProperties properties) {
            if (FontSize.HasValue) properties.FontSize = PowerPointTextDefaults.ToDrawingFontSize(FontSize.Value.ToPoints(), nameof(FontSize));
            if (FontName != null) PowerPointTextPropertyFormatting.SetFontName(properties, FontName);
            if (Color != null) PowerPointTextPropertyFormatting.SetColor(properties, Color);
            if (Bold.HasValue) properties.Bold = Bold.Value;
            if (Italic.HasValue) properties.Italic = Italic.Value;
            if (UnderlineStyle.HasValue) properties.Underline = UnderlineStyle.Value.ToOpenXml();
            else if (Underline.HasValue) properties.Underline = Underline.Value ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None;
            if (Strikethrough.HasValue) properties.Strike = Strikethrough.Value ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;
            if (HighlightColor != null) PowerPointTextPropertyFormatting.SetHighlightColor(properties, HighlightColor);
        }
    }
}
