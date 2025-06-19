using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Turns bold formatting on or off for this paragraph.
        /// </summary>
        public WordParagraph SetBold(bool isBold = true) {
            this.Bold = isBold;
            return this;
        }
        /// <summary>
        /// Sets italic formatting for the paragraph text.
        /// </summary>
        public WordParagraph SetItalic(bool isItalic = true) {
            this.Italic = isItalic;
            return this;
        }
        /// <summary>
        /// Applies the specified underline style to the paragraph.
        /// </summary>
        public WordParagraph SetUnderline(UnderlineValues underline) {
            this.Underline = underline;
            return this;
        }
        /// <summary>
        /// Sets the spacing after the paragraph in twentieths of a point.
        /// </summary>
        public WordParagraph SetSpacing(int spacing) {
            this.Spacing = spacing;
            return this;
        }
        /// <summary>
        /// Toggles single strike-through formatting for the paragraph.
        /// </summary>
        public WordParagraph SetStrike(bool isStrike = true) {
            this.Strike = isStrike;
            return this;
        }
        /// <summary>
        /// Toggles double strike-through formatting for the paragraph.
        /// </summary>
        public WordParagraph SetDoubleStrike(bool isDoubleStrike = true) {
            this.DoubleStrike = isDoubleStrike;
            return this;
        }
        /// <summary>
        /// Sets the font size of the paragraph text in points.
        /// </summary>
        public WordParagraph SetFontSize(int fontSize) {
            this.FontSize = fontSize;
            return this;
        }
        /// <summary>
        /// Sets the font family for the paragraph text.
        /// </summary>
        public WordParagraph SetFontFamily(string fontFamily) {
            this.FontFamily = fontFamily;
            return this;
        }
        /// <summary>
        /// Changes the text color using a hex string.
        /// </summary>
        public WordParagraph SetColorHex(string color) {
            this.ColorHex = color;
            return this;
        }
        /// <summary>
        /// Changes the text color using an <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        public WordParagraph SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color;
            return this;
        }
        /// <summary>
        /// Sets the paragraph alignment.
        /// </summary>
        public WordParagraph SetAlignment(JustificationValues alignment) {
            this.ParagraphAlignment = alignment;
            return this;
        }

        /// <summary>
        /// Applies a highlight color to the paragraph text.
        /// </summary>
        public WordParagraph SetHighlight(HighlightColorValues highlight) {
            this.Highlight = highlight;
            return this;
        }
        /// <summary>
        /// Sets the capitalization style for the paragraph text.
        /// </summary>
        public WordParagraph SetCapsStyle(CapsStyle capsStyle) {
            this.CapsStyle = capsStyle;
            return this;
        }
        /// <summary>
        /// Replaces the paragraph text with the specified value.
        /// </summary>
        public WordParagraph SetText(string text) {
            this.Text = text;
            return this;
        }
        /// <summary>
        /// Applies one of the predefined paragraph styles.
        /// </summary>
        public WordParagraph SetStyle(WordParagraphStyles style) {
            this.Style = style;
            return this;
        }


        /// <summary>
        /// Applies a paragraph style using its style identifier string.
        /// </summary>
        public WordParagraph SetStyleId(string styleId) {
            //Todo Check the styleId exist
            if (!string.IsNullOrEmpty(styleId)) {
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                if (_paragraphProperties.ParagraphStyleId == null) {
                    _paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
                }
                _paragraphProperties.ParagraphStyleId.Val = styleId;
            }
            return this;
        }

        /// <summary>
        /// Set the vertical text alignment
        /// </summary>
        /// <param name="verticalPositionValue"></param>
        /// <returns></returns>
        public WordParagraph SetVerticalTextAlignment(VerticalPositionValues? verticalPositionValue) {
            VerticalTextAlignment = verticalPositionValue;
            return this;
        }

        /// <summary>
        /// Set the text as subscript
        /// </summary>
        /// <returns></returns>
        public WordParagraph SetSubScript() {
            VerticalTextAlignment = VerticalPositionValues.Subscript;
            return this;
        }

        /// <summary>
        /// Set the text as superscript
        /// </summary>
        /// <returns></returns>
        public WordParagraph SetSuperScript() {
            VerticalTextAlignment = VerticalPositionValues.Superscript;
            return this;
        }
    }
}
