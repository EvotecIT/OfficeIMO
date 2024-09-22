using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        public WordParagraph SetBold(bool isBold = true) {
            this.Bold = isBold;
            return this;
        }
        public WordParagraph SetItalic(bool isItalic = true) {
            this.Italic = isItalic;
            return this;
        }
        public WordParagraph SetUnderline(UnderlineValues underline) {
            this.Underline = underline;
            return this;
        }
        public WordParagraph SetSpacing(int spacing) {
            this.Spacing = spacing;
            return this;
        }
        public WordParagraph SetStrike(bool isStrike = true) {
            this.Strike = isStrike;
            return this;
        }
        public WordParagraph SetDoubleStrike(bool isDoubleStrike = true) {
            this.DoubleStrike = isDoubleStrike;
            return this;
        }
        public WordParagraph SetFontSize(int fontSize) {
            this.FontSize = fontSize;
            return this;
        }
        public WordParagraph SetFontFamily(string fontFamily) {
            this.FontFamily = fontFamily;
            return this;
        }
        public WordParagraph SetColorHex(string color) {
            this.ColorHex = color;
            return this;
        }
        public WordParagraph SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color;
            return this;
        }
        public WordParagraph SetAlignment(JustificationValues alignment) {
            this.ParagraphAlignment = alignment;
            return this;
        }

        public WordParagraph SetHighlight(HighlightColorValues highlight) {
            this.Highlight = highlight;
            return this;
        }
        public WordParagraph SetCapsStyle(CapsStyle capsStyle) {
            this.CapsStyle = capsStyle;
            return this;
        }
        public WordParagraph SetText(string text) {
            this.Text = text;
            return this;
        }
        public WordParagraph SetStyle(WordParagraphStyles style) {
            this.Style = style;
            return this;
        }


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
