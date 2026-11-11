namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Splits this document into one new document per section.
        /// </summary>
        /// <remarks>
        /// Each piece keeps its section's page setup, headers and footers (including those it inherits from earlier
        /// sections), and only the styles, lists, images, hyperlinks, footnotes, endnotes and comments its content uses.
        /// Document settings, theme, fonts and properties are copied from this document. Bookmarks, comment ranges and
        /// fields that cross a piece boundary are trimmed so every piece is well formed. This document is not changed,
        /// but pending edits are flushed as they would be by <see cref="ToBytes(WordSaveOptions?)"/>.
        /// </remarks>
        /// <returns>The pieces in document order. Each is detached and must be saved with an explicit destination.</returns>
        public IReadOnlyList<WordDocument> SplitBySection() {
            return WordDocumentSplit.Split(this, headingLevel: null);
        }

        /// <summary>
        /// Splits this document into one new document per heading, starting a piece at every top-level paragraph whose
        /// heading level is <paramref name="level"/> or higher (for example every <c>Heading 1</c> when the level is 1).
        /// </summary>
        /// <remarks>
        /// Headings are paragraphs whose outline level, set on the paragraph or its style chain, or whose
        /// <c>heading N</c> style name, gives a level from 1 to 9. Content before the first heading becomes its own piece
        /// when it has text, a table or a picture. A piece that crosses a section break keeps it, and each piece keeps
        /// the page setup, headers and footers in effect where its content sits, together with only the styles, lists,
        /// images, hyperlinks, notes and comments it uses. This document is not changed, but pending edits are flushed as
        /// they would be by <see cref="ToBytes(WordSaveOptions?)"/>.
        /// </remarks>
        /// <param name="level">The deepest heading level that starts a new piece, from 1 to 9.</param>
        /// <returns>The pieces in document order. Each is detached and must be saved with an explicit destination.</returns>
        public IReadOnlyList<WordDocument> SplitByHeading(int level = 1) {
            if (level < 1 || level > 9) throw new ArgumentOutOfRangeException(nameof(level), level, "The heading level must be between 1 and 9.");
            return WordDocumentSplit.Split(this, level);
        }
    }
}
