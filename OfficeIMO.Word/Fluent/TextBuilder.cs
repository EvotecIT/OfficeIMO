using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Provides helpers for styling text fragments.
    /// </summary>
    public class TextBuilder {
        private readonly WordParagraph? _paragraph;

        internal TextBuilder(WordParagraph paragraph) {
            _paragraph = paragraph;
        }

        /// <summary>
        /// Gets the paragraph associated with this builder.
        /// </summary>
        public WordParagraph? Paragraph => _paragraph;

        /// <summary>
        /// Applies bold formatting to the current run.
        /// </summary>
        public TextBuilder BoldOn() {
            _paragraph?.SetBold();
            return this;
        }

        /// <summary>
        /// Applies italic formatting to the current run.
        /// </summary>
        public TextBuilder ItalicOn() {
            _paragraph?.SetItalic();
            return this;
        }

        /// <summary>
        /// Sets the text color using a hexadecimal value.
        /// </summary>
        /// <param name="hex">Color in hexadecimal format.</param>
        public TextBuilder Color(string hex) {
            if (hex.StartsWith("#")) {
                hex = hex.Substring(1);
            }
            _paragraph?.SetColorHex("#" + hex);
            return this;
        }

        /// <summary>
        /// Applies underline formatting to the current run.
        /// </summary>
        /// <param name="underline">Underline style.</param>
        public TextBuilder Underline(UnderlineValues underline) {
            _paragraph?.SetUnderline(underline);
            return this;
        }

        /// <summary>
        /// Applies single strikethrough formatting to the current run.
        /// </summary>
        /// <param name="isStrike">Whether to apply strikethrough.</param>
        public TextBuilder Strike(bool isStrike = true) {
            _paragraph?.SetStrike(isStrike);
            return this;
        }

        /// <summary>
        /// Applies double strikethrough formatting to the current run.
        /// </summary>
        /// <param name="isDoubleStrike">Whether to apply double strikethrough.</param>
        public TextBuilder DoubleStrike(bool isDoubleStrike = true) {
            _paragraph?.SetDoubleStrike(isDoubleStrike);
            return this;
        }

        /// <summary>
        /// Sets the font size for the current run.
        /// </summary>
        /// <param name="fontSize">Font size in points.</param>
        public TextBuilder FontSize(int fontSize) {
            _paragraph?.SetFontSize(fontSize);
            return this;
        }

        /// <summary>
        /// Sets the font family for the current run.
        /// </summary>
        /// <param name="fontFamily">Font family name.</param>
        public TextBuilder FontFamily(string fontFamily) {
            _paragraph?.SetFontFamily(fontFamily);
            return this;
        }

        /// <summary>
        /// Applies a highlight color to the current run.
        /// </summary>
        /// <param name="highlight">Highlight color.</param>
        public TextBuilder Highlight(HighlightColorValues highlight) {
            _paragraph?.SetHighlight(highlight);
            return this;
        }

        /// <summary>
        /// Sets the text as subscript.
        /// </summary>
        public TextBuilder SubScript() {
            _paragraph?.SetSubScript();
            return this;
        }

        /// <summary>
        /// Sets the text as superscript.
        /// </summary>
        public TextBuilder SuperScript() {
            _paragraph?.SetSuperScript();
            return this;
        }

        /// <summary>
        /// Applies capitalization style to the current run.
        /// </summary>
        /// <param name="capsStyle">Capitalization style.</param>
        public TextBuilder CapsStyle(CapsStyle capsStyle) {
            _paragraph?.SetCapsStyle(capsStyle);
            return this;
        }

        /// <summary>
        /// Applies outline formatting to the current run.
        /// </summary>
        public TextBuilder Outline() {
            _paragraph?.SetOutline();
            return this;
        }

        /// <summary>
        /// Applies shadow formatting to the current run.
        /// </summary>
        public TextBuilder Shadow() {
            _paragraph?.SetShadow();
            return this;
        }

        /// <summary>
        /// Applies emboss formatting to the current run.
        /// </summary>
        public TextBuilder Emboss() {
            _paragraph?.SetEmboss();
            return this;
        }

        /// <summary>
        /// Applies small caps formatting to the current run.
        /// </summary>
        public TextBuilder SmallCaps() {
            _paragraph?.SetSmallCaps();
            return this;
        }
    }
}