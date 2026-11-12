namespace OfficeIMO.Word {
    /// <summary>
    /// Controls how <see cref="WordDocument.AppendDocument(WordDocument, WordDocumentMergeOptions?)"/> reconciles styles and page setup.
    /// </summary>
    public sealed class WordDocumentMergeOptions {
        /// <summary>
        /// Gets or sets what happens when a style used by the appended content matches a destination style by name
        /// (or by identifier) but is defined differently. The default, <see cref="WordStyleConflictMode.Rename"/>,
        /// keeps the appended content looking as it did in the source.
        /// </summary>
        /// <remarks>
        /// <see cref="WordStyleConflictMode.KeepTarget"/> formats the appended content with the destination definition.
        /// <see cref="WordStyleConflictMode.Overwrite"/> replaces the destination definition with the source one, so
        /// destination content using that style changes too. <see cref="WordStyleConflictMode.Rename"/> copies the
        /// source style under a new identifier and name (for example <c>Heading1_0</c> / <c>heading 1_0</c>) and
        /// repoints the appended content to the copy. Matching styles with identical definitions are always shared,
        /// except styles that number paragraphs with a list, which are treated as different. Source styles the
        /// destination lacks are copied in every mode.
        /// </remarks>
        public WordStyleConflictMode StyleConflictMode { get; set; } = WordStyleConflictMode.Rename;

        /// <summary>
        /// Gets or sets how the page setup of the appended document is applied.
        /// The default starts the appended content in its own section with the source page setup.
        /// </summary>
        public WordMergeSectionMode SectionMode { get; set; } = WordMergeSectionMode.KeepSourceSections;
    }

    /// <summary>
    /// Specifies how section page setup is handled when one document is appended to another.
    /// </summary>
    public enum WordMergeSectionMode {
        /// <summary>
        /// The destination's final section is closed with a section break and the appended content keeps
        /// its own sections, page setup, headers and footers.
        /// </summary>
        KeepSourceSections,

        /// <summary>
        /// The appended content continues in the destination's final section; source section breaks and page setup are dropped.
        /// </summary>
        UseDestinationSections
    }
}
