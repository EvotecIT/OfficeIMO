namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Identifies non-compound legacy DOC feature metadata preserved during import.
    /// </summary>
    public enum LegacyDocPreservedFeatureKind {
        /// <summary>
        /// Picture payloads were indicated by the FIB and remain preserve-only import metadata.
        /// </summary>
        Picture,

        /// <summary>
        /// Revision tracking metadata was indicated by document properties and remains preserve-only import metadata.
        /// </summary>
        RevisionTracking,

        /// <summary>
        /// A bookmark range exists outside the currently projected document stories.
        /// </summary>
        Bookmark
    }
}
