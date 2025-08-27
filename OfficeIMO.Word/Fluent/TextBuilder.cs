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
    }
}
