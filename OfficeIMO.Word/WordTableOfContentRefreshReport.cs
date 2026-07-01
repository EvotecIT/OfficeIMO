namespace OfficeIMO.Word {
    /// <summary>
    /// Describes a generated table-of-contents entry.
    /// </summary>
    public sealed class WordTableOfContentEntry {
        internal WordTableOfContentEntry(string text, int level, int pageNumber, string bookmarkName) {
            Text = text;
            Level = level;
            PageNumber = pageNumber;
            BookmarkName = bookmarkName;
        }

        /// <summary>Gets the heading text used for the TOC entry.</summary>
        public string Text { get; }

        /// <summary>Gets the heading level from 1 to 9.</summary>
        public int Level { get; }

        /// <summary>Gets the page number estimated from explicit page breaks.</summary>
        public int PageNumber { get; }

        /// <summary>Gets the bookmark anchor used by the generated internal hyperlink.</summary>
        public string BookmarkName { get; }
    }

    /// <summary>
    /// Reports the result of an OfficeIMO-generated table-of-contents refresh.
    /// </summary>
    public sealed class WordTableOfContentRefreshReport {
        internal WordTableOfContentRefreshReport(IReadOnlyList<WordTableOfContentEntry> entries, int skippedHeadingCount, string pageNumberMode) {
            Entries = entries.ToArray();
            SkippedHeadingCount = skippedHeadingCount;
            PageNumberMode = pageNumberMode;
        }

        /// <summary>Gets generated entries in document order.</summary>
        public IReadOnlyList<WordTableOfContentEntry> Entries { get; }

        /// <summary>Gets the number of headings ignored because they were outside the configured TOC level range.</summary>
        public int SkippedHeadingCount { get; }

        /// <summary>Gets a short description of how page numbers were calculated.</summary>
        public string PageNumberMode { get; }

        /// <summary>Gets the number of generated entries.</summary>
        public int EntryCount => Entries.Count;
    }
}
