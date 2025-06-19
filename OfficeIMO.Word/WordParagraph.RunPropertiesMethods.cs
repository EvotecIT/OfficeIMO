using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Executes the SetBold operation.
        /// </summary>
        public WordParagraph SetBold(bool isBold = true) {
            this.Bold = isBold;
            return this;
        }
        /// <summary>
        /// Executes the SetItalic operation.
        /// </summary>
        public WordParagraph SetItalic(bool isItalic = true) {
            this.Italic = isItalic;
            return this;
        }
        /// <summary>
        /// Executes the SetUnderline operation.
        /// </summary>
        public WordParagraph SetUnderline(UnderlineValues underline) {
            this.Underline = underline;
            return this;
        }
        /// <summary>
        /// Executes the SetSpacing operation.
        /// </summary>
        public WordParagraph SetSpacing(int spacing) {
            this.Spacing = spacing;
            return this;
        }
        /// <summary>
        /// Executes the SetStrike operation.
        /// </summary>
        public WordParagraph SetStrike(bool isStrike = true) {
            this.Strike = isStrike;
            return this;
        }
        /// <summary>
        /// Executes the SetDoubleStrike operation.
        /// </summary>
        public WordParagraph SetDoubleStrike(bool isDoubleStrike = true) {
            this.DoubleStrike = isDoubleStrike;
            return this;
        }
        /// <summary>
        /// Executes the SetFontSize operation.
        /// </summary>
        public WordParagraph SetFontSize(int fontSize) {
            this.FontSize = fontSize;
            return this;
        }
        /// <summary>
        /// Executes the SetFontFamily operation.
        /// </summary>
        public WordParagraph SetFontFamily(string fontFamily) {
            this.FontFamily = fontFamily;
            return this;
        }
        /// <summary>
        /// Executes the SetColorHex operation.
        /// </summary>
        public WordParagraph SetColorHex(string color) {
            this.ColorHex = color;
            return this;
        }
        /// <summary>
        /// Executes the SetColor operation.
        /// </summary>
        public WordParagraph SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color;
            return this;
        }
        /// <summary>
        /// Executes the SetAlignment operation.
        /// </summary>
        public WordParagraph SetAlignment(JustificationValues alignment) {
            this.ParagraphAlignment = alignment;
            return this;
        }

        /// <summary>
        /// Executes the SetHighlight operation.
        /// </summary>
        public WordParagraph SetHighlight(HighlightColorValues highlight) {
            this.Highlight = highlight;
            return this;
        }
        /// <summary>
        /// Executes the SetCapsStyle operation.
        /// </summary>
        public WordParagraph SetCapsStyle(CapsStyle capsStyle) {
            this.CapsStyle = capsStyle;
            return this;
        }
        /// <summary>
        /// Executes the SetText operation.
        /// </summary>
        public WordParagraph SetText(string text) {
            this.Text = text;
            return this;
        }
        /// <summary>
        /// Executes the SetStyle operation.
        /// </summary>
        public WordParagraph SetStyle(WordParagraphStyles style) {
            this.Style = style;
            return this;
        }


        /// <summary>
        /// Executes the SetStyleId operation.
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
