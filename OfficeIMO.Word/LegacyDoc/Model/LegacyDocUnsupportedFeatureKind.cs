namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Identifies unsupported or preserve-only legacy DOC features discovered during import.
    /// </summary>
    public enum LegacyDocUnsupportedFeatureKind {
        /// <summary>
        /// VBA project storage was discovered in the OLE compound container.
        /// </summary>
        VbaProject,

        /// <summary>
        /// Embedded OLE object storage was discovered in the OLE compound container.
        /// </summary>
        OleObject,

        /// <summary>
        /// Header or footer story text exists but is not projected yet.
        /// </summary>
        HeaderFooter,

        /// <summary>
        /// Footnote story text exists but is not projected yet.
        /// </summary>
        Footnote,

        /// <summary>
        /// Endnote story text exists but is not projected yet.
        /// </summary>
        Endnote,

        /// <summary>
        /// Comment or annotation story text exists but is not projected yet.
        /// </summary>
        Comment,

        /// <summary>
        /// Text box story text exists but is not projected yet.
        /// </summary>
        TextBox
    }
}
