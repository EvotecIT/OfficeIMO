using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Defines text formatting that can be applied to runs or paragraphs.
    /// </summary>
    public readonly struct PowerPointTextStyle {
        /// <summary>
        /// Creates a new text style instance.
        /// </summary>
        public PowerPointTextStyle(int? fontSize = null, string? fontName = null, string? color = null,
            bool? bold = null, bool? italic = null, bool? underline = null) {
            FontSize = fontSize;
            FontName = fontName;
            Color = color;
            Bold = bold;
            Italic = italic;
            Underline = underline;
        }

        /// <summary>
        /// Font size in points.
        /// </summary>
        public int? FontSize { get; }

        /// <summary>
        /// Font name (Latin).
        /// </summary>
        public string? FontName { get; }

        /// <summary>
        /// Text color in hex (e.g. "1F4E79").
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

        /// <summary>
        /// A preset for typical slide titles.
        /// </summary>
        public static PowerPointTextStyle Title => new(fontSize: 32, bold: true);

        /// <summary>
        /// A preset for subtitles.
        /// </summary>
        public static PowerPointTextStyle Subtitle => new(fontSize: 24);

        /// <summary>
        /// A preset for body text.
        /// </summary>
        public static PowerPointTextStyle Body => new(fontSize: 18);

        /// <summary>
        /// A preset for captions or footnotes.
        /// </summary>
        public static PowerPointTextStyle Caption => new(fontSize: 12);

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
            if (Underline != null) {
                run.Underline = Underline.Value;
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
            if (Underline != null) {
                paragraph.SetUnderline(Underline.Value);
            }
        }

        /// <summary>
        /// Returns a copy with a new font size.
        /// </summary>
        public PowerPointTextStyle WithFontSize(int? fontSize) {
            return new PowerPointTextStyle(fontSize, FontName, Color, Bold, Italic, Underline);
        }

        /// <summary>
        /// Returns a copy with a new font name.
        /// </summary>
        public PowerPointTextStyle WithFontName(string? fontName) {
            return new PowerPointTextStyle(FontSize, fontName, Color, Bold, Italic, Underline);
        }

        /// <summary>
        /// Returns a copy with a new color.
        /// </summary>
        public PowerPointTextStyle WithColor(string? color) {
            return new PowerPointTextStyle(FontSize, FontName, color, Bold, Italic, Underline);
        }

        /// <summary>
        /// Returns a copy with bold formatting updated.
        /// </summary>
        public PowerPointTextStyle WithBold(bool? bold) {
            return new PowerPointTextStyle(FontSize, FontName, Color, bold, Italic, Underline);
        }

        /// <summary>
        /// Returns a copy with italic formatting updated.
        /// </summary>
        public PowerPointTextStyle WithItalic(bool? italic) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, italic, Underline);
        }

        /// <summary>
        /// Returns a copy with underline formatting updated.
        /// </summary>
        public PowerPointTextStyle WithUnderline(bool? underline) {
            return new PowerPointTextStyle(FontSize, FontName, Color, Bold, Italic, underline);
        }
    }
}
