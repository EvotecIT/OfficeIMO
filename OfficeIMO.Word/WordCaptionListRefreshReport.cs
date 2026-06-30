namespace OfficeIMO.Word {
    /// <summary>
    /// Describes a generated list-of-figures or list-of-tables entry.
    /// </summary>
    public sealed class WordCaptionListEntry {
        internal WordCaptionListEntry(string sequenceIdentifier, string text, int pageNumber, string bookmarkName) {
            SequenceIdentifier = sequenceIdentifier;
            Text = text;
            PageNumber = pageNumber;
            BookmarkName = bookmarkName;
        }

        /// <summary>Gets the SEQ identifier used by the caption, for example Figure or Table.</summary>
        public string SequenceIdentifier { get; }

        /// <summary>Gets the visible caption text used for the generated entry.</summary>
        public string Text { get; }

        /// <summary>Gets the page number estimated from explicit page breaks and section starts.</summary>
        public int PageNumber { get; }

        /// <summary>Gets the bookmark anchor used by the generated internal hyperlink.</summary>
        public string BookmarkName { get; }
    }

    /// <summary>
    /// Reports the result of an OfficeIMO-generated caption-list refresh.
    /// </summary>
    public sealed class WordCaptionListRefreshReport {
        internal WordCaptionListRefreshReport(
            string sequenceIdentifier,
            IReadOnlyList<WordCaptionListEntry> entries,
            int skippedCaptionCount,
            string pageNumberMode) {
            SequenceIdentifier = sequenceIdentifier;
            Entries = entries.ToArray();
            SkippedCaptionCount = skippedCaptionCount;
            PageNumberMode = pageNumberMode;
        }

        /// <summary>Gets the SEQ identifier used to collect captions.</summary>
        public string SequenceIdentifier { get; }

        /// <summary>Gets generated entries in document order.</summary>
        public IReadOnlyList<WordCaptionListEntry> Entries { get; }

        /// <summary>Gets the number of matching captions ignored because they had no visible text.</summary>
        public int SkippedCaptionCount { get; }

        /// <summary>Gets a short description of how page numbers were calculated.</summary>
        public string PageNumberMode { get; }

        /// <summary>Gets the number of generated entries.</summary>
        public int EntryCount => Entries.Count;
    }
}
