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
        OleObject
    }
}
