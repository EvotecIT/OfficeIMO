namespace OfficeIMO.Word {
    /// <summary>
    /// Describes a generated index entry.
    /// </summary>
    public sealed class WordIndexEntry {
        internal WordIndexEntry(string term, IReadOnlyList<string> subterms, IReadOnlyList<int> pageNumbers, IReadOnlyList<string> pageReferences, string? crossReferenceText, string? entryType, string pageReferenceSeparator = ", ") {
            Term = term;
            Subterms = subterms.ToArray();
            PageNumbers = pageNumbers.ToArray();
            PageReferences = pageReferences.ToArray();
            CrossReferenceText = crossReferenceText;
            EntryType = entryType;
            PageReferenceSeparator = pageReferenceSeparator;
        }

        /// <summary>Gets the main index term.</summary>
        public string Term { get; }

        /// <summary>Gets the optional first-level subentry. For deeper entries, use <see cref="Subterms"/>.</summary>
        public string? Subterm => Subterms.Count > 0 ? Subterms[0] : null;

        /// <summary>Gets all subentry levels under the main term.</summary>
        public IReadOnlyList<string> Subterms { get; }

        /// <summary>Gets the complete index path, including the main term and every subentry level.</summary>
        public IReadOnlyList<string> Path => new[] { Term }.Concat(Subterms).ToArray();

        /// <summary>Gets the cross-reference text for entries such as "See another term".</summary>
        public string? CrossReferenceText { get; }

        /// <summary>Gets the optional Word index entry type from the <c>XE \f</c> switch.</summary>
        public string? EntryType { get; }

        /// <summary>Gets whether this entry renders as a cross-reference instead of page numbers.</summary>
        public bool IsCrossReference => !string.IsNullOrWhiteSpace(CrossReferenceText);

        /// <summary>Gets estimated page numbers collected for the entry.</summary>
        public IReadOnlyList<int> PageNumbers { get; }

        /// <summary>Gets estimated page references collected for the entry, including bookmark page ranges such as <c>1-2</c>.</summary>
        public IReadOnlyList<string> PageReferences { get; }

        /// <summary>Gets page numbers and page ranges formatted for display.</summary>
        public string PageNumbersText => string.Join(PageReferenceSeparator, PageReferences);

        private string PageReferenceSeparator { get; }
    }

    /// <summary>
    /// Reports the result of an OfficeIMO-generated index refresh.
    /// </summary>
    public sealed class WordIndexRefreshReport {
        internal WordIndexRefreshReport(IReadOnlyList<WordIndexEntry> entries, int skippedEntryCount, string pageNumberMode, int? columnCount) {
            Entries = entries.ToArray();
            SkippedEntryCount = skippedEntryCount;
            PageNumberMode = pageNumberMode;
            ColumnCount = columnCount;
        }

        /// <summary>Gets generated entries in sorted index order.</summary>
        public IReadOnlyList<WordIndexEntry> Entries { get; }

        /// <summary>Gets the number of XE fields ignored because they were unsupported or malformed.</summary>
        public int SkippedEntryCount { get; }

        /// <summary>Gets a short description of how page numbers were calculated.</summary>
        public string PageNumberMode { get; }

        /// <summary>Gets the bounded column count requested by an imported <c>INDEX \c</c> switch, when present.</summary>
        public int? ColumnCount { get; }

        /// <summary>Gets the number of generated entries.</summary>
        public int EntryCount => Entries.Count;
    }
}
