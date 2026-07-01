namespace OfficeIMO.Word {
    /// <summary>
    /// Reports the result of marking a document from Word index concordance entries.
    /// </summary>
    public sealed class WordIndexConcordanceReport {
        internal WordIndexConcordanceReport(
            IReadOnlyList<WordIndexConcordanceEntry> entries,
            int markedEntryCount,
            int matchedParagraphCount,
            int skippedEntryCount,
            bool matchCase,
            bool matchWholeWord) {
            Entries = entries.ToArray();
            MarkedEntryCount = markedEntryCount;
            MatchedParagraphCount = matchedParagraphCount;
            SkippedEntryCount = skippedEntryCount;
            MatchCase = matchCase;
            MatchWholeWord = matchWholeWord;
        }

        /// <summary>Gets valid concordance entries that were considered for marking.</summary>
        public IReadOnlyList<WordIndexConcordanceEntry> Entries { get; }

        /// <summary>Gets the number of valid concordance entries considered for marking.</summary>
        public int ConcordanceEntryCount => Entries.Count;

        /// <summary>Gets the number of hidden <c>XE</c> fields inserted into the target document.</summary>
        public int MarkedEntryCount { get; }

        /// <summary>Gets the number of distinct paragraphs that received one or more marks.</summary>
        public int MatchedParagraphCount { get; }

        /// <summary>Gets the number of blank, invalid, unsafe, or duplicate concordance entries skipped.</summary>
        public int SkippedEntryCount { get; }

        /// <summary>Gets whether text matching was case-sensitive.</summary>
        public bool MatchCase { get; }

        /// <summary>Gets whether matches required non-letter/digit boundaries around the search text.</summary>
        public bool MatchWholeWord { get; }
    }
}
