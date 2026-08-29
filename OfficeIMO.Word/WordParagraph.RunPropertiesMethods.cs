using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using M = DocumentFormat.OpenXml.Math;

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
        public WordParagraph SetUnderline(WordUnderlineStyle underline) {
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
            this.CapsStyle = isSmallCaps ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
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
        /// Sets the language for the text run.
        /// </summary>
        /// <param name="language">Language tag, such as <c>en-US</c> or <c>pl-PL</c>.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetLanguage(string language) {
            this.Language = language;
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
        /// Sets the text color using <see cref="OfficeIMO.Drawing.OfficeColor"/>.
        /// </summary>
        /// <param name="color">The color to apply.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetColor(OfficeIMO.Drawing.OfficeColor color) {
            this.Color = color;
            return this;
        }
        /// <summary>
        /// Sets the paragraph alignment.
        /// </summary>
        /// <param name="alignment">Alignment value.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetAlignment(WordParagraphAlignment alignment) {
            this.ParagraphAlignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets the highlight color for the text.
        /// </summary>
        /// <param name="highlight">Highlight color.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetHighlight(WordHighlightColor highlight) {
            this.Highlight = highlight;
            return this;
        }
        /// <summary>
        /// Sets the capitalization style for the text.
        /// </summary>
        /// <param name="capsStyle">Capitalization style.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph SetCapsStyle(WordCapsStyle capsStyle) {
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
        /// Changes the stored run text casing while preserving run formatting.
        /// </summary>
        /// <param name="textCase">Casing transformation to apply.</param>
        /// <param name="culture">Culture used for casing. The current culture is used when omitted.</param>
        /// <returns>The current paragraph instance.</returns>
        public WordParagraph TransformTextCase(OfficeIMO.Drawing.OfficeTextCase textCase, CultureInfo? culture = null) {
            WordEquation? equation = Equation;
            if (equation?.Representation == WordEquationRepresentation.Omml) {
                M.Text[] textNodes = equation.MathElement?.Descendants<M.Text>().ToArray() ?? Array.Empty<M.Text>();
                IReadOnlyList<string> transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.ApplySegments(
                    textNodes.Select(node => node.Text ?? string.Empty).ToArray(), textCase, culture);
                for (int index = 0; index < textNodes.Length; index++) textNodes[index].Text = transformed[index];
                return this;
            }

            IReadOnlyList<OpenXmlElement> roots = GetTextCaseTransformationRoots();
            var segmentNodes = new List<Text?>();
            var segments = new List<string>();
            foreach (OpenXmlElement root in roots) {
                foreach (OpenXmlElement element in root.Descendants()) {
                    switch (element) {
                        case Text textNode:
                            segmentNodes.Add(textNode);
                            segments.Add(textNode.Text ?? string.Empty);
                            break;
                        case TabChar:
                            segmentNodes.Add(null);
                            segments.Add("\t");
                            break;
                        case Break breakNode:
                            segmentNodes.Add(null);
                            segments.Add(IsTextWrappingBreak(breakNode) ? "\n" : "\u2028");
                            break;
                        case CarriageReturn:
                            segmentNodes.Add(null);
                            segments.Add("\n");
                            break;
                    }
                }
            }
            if (segmentNodes.Any(node => node != null)) {
                IReadOnlyList<string> transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.ApplySegments(
                    segments, textCase, culture);
                for (int index = 0; index < segmentNodes.Count; index++) {
                    if (segmentNodes[index] != null) segmentNodes[index]!.Text = transformed[index];
                }
                return this;
            }
            Text = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(Text, textCase, culture);
            return this;
        }

        private IReadOnlyList<OpenXmlElement> GetTextCaseTransformationRoots() {
            if (_run != null) return new OpenXmlElement[] { _run };
            if (_hyperlink != null) return new OpenXmlElement[] { _hyperlink };
            if (_simpleField != null) return new OpenXmlElement[] { _simpleField };
            if (_runs != null && _runs.Count > 0) return GetComplexFieldResultRuns(_runs).Cast<OpenXmlElement>().ToArray();
            if (_stdRun?.SdtContentRun != null) return new OpenXmlElement[] { _stdRun.SdtContentRun };
            return Array.Empty<OpenXmlElement>();
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
        /// <example>
        /// <code><![CDATA[
        /// Style style = new() {
        ///     Type = StyleValues.Paragraph,
        ///     StyleId = "MyStyle"
        /// };
        ///
        /// WordParagraphStyle.RegisterCustomStyle("MyStyle", style);
        /// document.AddParagraph("Hello world").SetStyleId("MyStyle");
        /// ]]></code>
        /// </example>
        public WordParagraph SetStyleId(string styleId) {
            //Todo Check the styleId exist
            if (!string.IsNullOrEmpty(styleId)) {
                _document?.EnsureStyleDefinitionsInitialized();
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
        public WordParagraph SetVerticalTextAlignment(WordVerticalTextPosition? verticalPositionValue) {
            VerticalTextAlignment = verticalPositionValue;
            return this;
        }

        /// <summary>
        /// Set the text as subscript
        /// </summary>
        /// <returns></returns>
        public WordParagraph SetSubScript() {
            VerticalTextAlignment = WordVerticalTextPosition.Subscript;
            return this;
        }

        /// <summary>
        /// Set the text as superscript
        /// </summary>
        /// <returns></returns>
        public WordParagraph SetSuperScript() {
            VerticalTextAlignment = WordVerticalTextPosition.Superscript;
            return this;
        }
    }
}
