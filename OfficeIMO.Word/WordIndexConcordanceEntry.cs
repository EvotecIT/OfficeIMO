namespace OfficeIMO.Word {
    /// <summary>
    /// Describes one Word index concordance mapping from document text to an index entry.
    /// </summary>
    public sealed class WordIndexConcordanceEntry {
        /// <summary>
        /// Creates a concordance mapping.
        /// </summary>
        /// <param name="searchText">Text to find in the target document.</param>
        /// <param name="indexText">Index entry text to write into an <c>XE</c> field. Use colons for subentries, for example <c>Policy:Alpha</c>.</param>
        public WordIndexConcordanceEntry(string searchText, string indexText) {
            if (string.IsNullOrWhiteSpace(searchText)) {
                throw new ArgumentException("Concordance search text cannot be empty.", nameof(searchText));
            }

            if (string.IsNullOrWhiteSpace(indexText)) {
                throw new ArgumentException("Concordance index text cannot be empty.", nameof(indexText));
            }

            SearchText = searchText.Trim();
            IndexText = indexText.Trim();
        }

        /// <summary>Gets the text to find in the target document.</summary>
        public string SearchText { get; }

        /// <summary>Gets the index entry text written into the hidden <c>XE</c> field.</summary>
        public string IndexText { get; }
    }
}
