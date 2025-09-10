using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Offers methods to modify run properties.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Sets the <see cref="WordParagraph.Bold"/> property.
        /// </summary>
        /// <param name="isBold">Whether the text should be bold.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetBold(bool isBold = true) {
            this.Bold = isBold;
            return this;
        }
        /// <summary>
        /// Sets the <see cref="WordParagraph.Italic"/> property.
        /// </summary>
        /// <param name="isItalic">Whether the text should be italic.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetItalic(bool isItalic = true) {
            this.Italic = isItalic;
            return this;
        }
        /// <summary>
        /// Sets the underline style for the text.
        /// </summary>
        /// <param name="underline">Underline style.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetUnderline(UnderlineValues underline) {
            this.Underline = underline;
            return this;
        }
        /// <summary>
        /// Sets the character spacing for the text.
        /// </summary>
        /// <param name="spacing">Spacing value in twentieths of a point.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetSpacing(int spacing) {
            this.Spacing = spacing;
            return this;
        }
        /// <summary>
        /// Enables or disables single strikethrough on the text.
        /// </summary>
        /// <param name="isStrike">Whether the text should be struck through.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetStrike(bool isStrike = true) {
            this.Strike = isStrike;
            return this;
        }
        /// <summary>
        /// Enables or disables double strikethrough on the text.
        /// </summary>
        /// <param name="isDoubleStrike">Whether the text should be double struck.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetDoubleStrike(bool isDoubleStrike = true) {
            this.DoubleStrike = isDoubleStrike;
            return this;
        }

        /// <summary>
        /// Enables or disables outline effect on the text.
        /// </summary>
        /// <param name="isOutline">Whether the text should be outlined.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetOutline(bool isOutline = true) {
            this.Outline = isOutline;
            return this;
        }

        /// <summary>
        /// Enables or disables shadow effect on the text.
        /// </summary>
        /// <param name="isShadow">Whether the text should have a shadow.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetShadow(bool isShadow = true) {
            this.Shadow = isShadow;
            return this;
        }

        /// <summary>
        /// Enables or disables emboss effect on the text.
        /// </summary>
        /// <param name="isEmboss">Whether the text should be embossed.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetEmboss(bool isEmboss = true) {
            this.Emboss = isEmboss;
            return this;
        }

        /// <summary>
        /// Enables or disables small caps formatting on the text.
        /// </summary>
        /// <param name="isSmallCaps">Whether the text should use small caps.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetSmallCaps(bool isSmallCaps = true) {
            this.CapsStyle = isSmallCaps ? CapsStyle.SmallCaps : CapsStyle.None;
            return this;
        }
        /// <summary>
        /// Sets the font size for the text.
        /// </summary>
        /// <param name="fontSize">Font size in points.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetFontSize(int fontSize) {
            this.FontSize = fontSize;
            return this;
        }
        /// <summary>
        /// Sets the font family for the text.
        /// </summary>
        /// <param name="fontFamily">Name of the font family.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetFontFamily(string fontFamily) {
            this.FontFamily = fontFamily;
            return this;
        }
        /// <summary>
        /// Sets the text color using a hexadecimal value.
        /// </summary>
        /// <param name="color">Color in hexadecimal format.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetColorHex(string color) {
            this.ColorHex = color;
            return this;
        }
        /// <summary>
        /// Sets the text color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        /// <param name="color">The color to apply.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color;
            return this;
        }
        /// <summary>
        /// Sets the paragraph alignment.
        /// </summary>
        /// <param name="alignment">Alignment value.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetAlignment(JustificationValues alignment) {
            this.ParagraphAlignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets the highlight color for the text.
        /// </summary>
        /// <param name="highlight">Highlight color.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetHighlight(HighlightColorValues highlight) {
            this.Highlight = highlight;
            return this;
        }
        /// <summary>
        /// Sets the capitalization style for the text.
        /// </summary>
        /// <param name="capsStyle">Capitalization style.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetCapsStyle(CapsStyle capsStyle) {
            this.CapsStyle = capsStyle;
            return this;
        }
        /// <summary>
        /// Sets the paragraph text.
        /// </summary>
        /// <param name="text">The text content.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetText(string text) {
            this.Text = text;
            return this;
        }
        /// <summary>
        /// Sets the paragraph style.
        /// </summary>
        /// <param name="style">The style to apply.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetStyle(WordParagraphStyles style) {
            this.Style = style;
            return this;
        }


        /// <summary>
        /// Sets the paragraph style by identifier.
        /// </summary>
        /// <param name="styleId">The style identifier.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetStyleId(string styleId) {
            //Todo Check the styleId exist
            if (!string.IsNullOrEmpty(styleId)) {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (props.ParagraphStyleId == null) {
                    props.ParagraphStyleId = new ParagraphStyleId();
                }
                props.ParagraphStyleId.Val = styleId;
            }
            return this;
        }

        /// <summary>
        /// Sets the character style for the run.
        /// </summary>
        /// <param name="style">Character style to apply.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetCharacterStyle(WordCharacterStyles style) {
            CharacterStyle = style;
            return this;
        }

        /// <summary>
        /// Sets the character style by identifier.
        /// </summary>
        /// <param name="styleId">The style identifier.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetCharacterStyleId(string styleId) {
            CharacterStyleId = styleId;
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
