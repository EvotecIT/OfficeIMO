namespace OfficeIMO.Word {
    /// <summary>
    /// Controls deterministic behavior used while refreshing supported Word fields.
    /// </summary>
    public sealed class WordFieldUpdateOptions {
        /// <summary>
        /// Gets an options instance that uses the current local date and time for DATE and TIME fields.
        /// </summary>
        public static WordFieldUpdateOptions Default => new WordFieldUpdateOptions();

        /// <summary>
        /// Gets or sets the date and time used for DATE and TIME field refresh.
        /// When unset, OfficeIMO uses the current local date and time.
        /// </summary>
        public DateTime? CurrentDateTime { get; set; }

        /// <summary>
        /// When true, refreshes only DATE and TIME fields and leaves all other field results unchanged.
        /// </summary>
        public bool DateTimeFieldsOnly { get; set; }

        /// <summary>
        /// When true, every table of contents is regenerated from the document headings as a complete TOC field
        /// result: one hyperlinked entry per heading with a <c>PAGEREF</c> page number. The entries are final, so the
        /// table is no longer flagged for Word to rebuild on open.
        /// </summary>
        public bool UpdateTablesOfContents { get; set; }

        /// <summary>
        /// When true, every table of contents is regenerated as with <see cref="UpdateTablesOfContents"/>, and only the
        /// fields inside the regenerated tables, their <c>PAGEREF</c> page numbers, are refreshed. Every other field
        /// keeps its result.
        /// </summary>
        /// <remarks>
        /// Use it after <see cref="WordDocument.AppendDocument(WordDocument, WordDocumentMergeOptions?)"/>. A full refresh
        /// evaluates fields such as <c>DOCPROPERTY</c>, <c>TITLE</c> and <c>AUTHOR</c> in the appended content against the
        /// destination's properties, so they would show the destination's values instead of their source's.
        /// Cannot be combined with <see cref="DateTimeFieldsOnly"/>.
        /// </remarks>
        public bool TablesOfContentsOnly { get; set; }

        /// <summary>
        /// When true, <c>PAGEREF</c> (including table-of-contents page numbers) and <c>NUMPAGES</c> fields use page
        /// numbers from OfficeIMO's own page layout, the engine behind PDF conversion, instead of counting explicit
        /// page and section breaks. This lays out a copy of the document, so it costs about as much as a PDF conversion.
        /// </summary>
        /// <remarks>Select the native engine with <c>WithPdfLayoutPagination()</c> from OfficeIMO.Word.Pdf.</remarks>
        public bool UseLayoutPagination { get; set; }

        /// <summary>The explicitly selected rendering owner; Word itself has no PDF dependency.</summary>
        internal Func<WordDocument, WordLayoutPagination>? LayoutPaginationProvider { get; set; }
    }
}
