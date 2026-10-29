namespace OfficeIMO.Word {
    /// <summary>
    /// Controls how <see cref="WordDocument.AppendDocument(WordDocument, WordDocumentMergeOptions?)"/> imports styles and page setup.
    /// </summary>
    public sealed class WordDocumentMergeOptions {
        /// <summary>
        /// Gets or sets how source styles that share a name or id with a destination style are imported.
        /// The default keeps the destination definitions.
        /// </summary>
        public WordMergeStyleMode StyleMode { get; set; } = WordMergeStyleMode.UseDestinationStyles;

        /// <summary>
        /// Gets or sets how the page setup of the appended document is applied.
        /// The default starts the appended content in its own section with the source page setup.
        /// </summary>
        public WordMergeSectionMode SectionMode { get; set; } = WordMergeSectionMode.KeepSourceSections;
    }

    /// <summary>
    /// Specifies how styles are reconciled when one document is appended to another.
    /// </summary>
    public enum WordMergeStyleMode {
        /// <summary>
        /// A source style that matches a destination style by name (or by id) uses the destination definition.
        /// Source styles missing from the destination are copied.
        /// </summary>
        UseDestinationStyles,

        /// <summary>
        /// A source style that matches a destination style but has a different definition is copied under a new
        /// unique id and name (for example <c>Heading1_0</c> / <c>heading 1_0</c>), and the appended content is
        /// repointed to the copy. Styles with identical definitions are shared.
        /// </summary>
        KeepDifferentStyles
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
