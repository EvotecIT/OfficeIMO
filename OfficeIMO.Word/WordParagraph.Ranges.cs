namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Returns a range covering all visible text of the paragraph that contains this element.</summary>
        /// <returns>The paragraph range.</returns>
        public WordRange GetRange() {
            int length = WordTextMap.Build(_paragraph).Text.Length;
            return new WordRange(_document, _paragraph, 0, _paragraph, length);
        }

        /// <summary>Returns a range of the visible text of the paragraph that contains this element.</summary>
        /// <param name="start">Zero-based offset from the start of the paragraph text, counting all of its runs.</param>
        /// <param name="length">Number of characters; zero gives an empty range that marks a position.</param>
        /// <returns>The range.</returns>
        /// <remarks>Use <see cref="WordRange.ExtendTo(WordRange)"/> to continue a range into a later paragraph.</remarks>
        public WordRange GetRange(int start, int length) {
            int textLength = WordTextMap.Build(_paragraph).Text.Length;
            if (start < 0 || start > textLength) throw new ArgumentOutOfRangeException(nameof(start));
            if (length < 0 || start + length > textLength) throw new ArgumentOutOfRangeException(nameof(length));
            return new WordRange(_document, _paragraph, start, _paragraph, start + length);
        }
    }
}
