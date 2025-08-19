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

        public WordParagraph? Paragraph => _paragraph;

        public TextBuilder Bold() {
            _paragraph?.SetBold();
            return this;
        }

        public TextBuilder Italic() {
            _paragraph?.SetItalic();
            return this;
        }

        public TextBuilder Color(string hex) {
            _paragraph?.SetColorHex(hex);
            return this;
        }
    }
}
