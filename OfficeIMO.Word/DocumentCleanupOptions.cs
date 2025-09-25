namespace OfficeIMO.Word {
    /// <summary>
    /// Specifies document cleanup operations.
    /// </summary>
    [Flags]
    public enum DocumentCleanupOptions {
        /// <summary>
        /// No cleanup.
        /// </summary>
        None = 0,
        /// <summary>
        /// Merge consecutive runs that share identical formatting.
        /// </summary>
        MergeIdenticalRuns = 1,
        /// <summary>
        /// Remove runs with no text content.
        /// </summary>
        RemoveEmptyRuns = 2,
        /// <summary>
        /// Remove run properties that have no children.
        /// </summary>
        RemoveRedundantRunProperties = 4,
        /// <summary>
        /// Delete paragraphs that end up with no runs.
        /// </summary>
        RemoveEmptyParagraphs = 8,
        /// <summary>
        /// Perform all cleanup operations.
        /// </summary>
        All = MergeIdenticalRuns | RemoveEmptyRuns | RemoveRedundantRunProperties | RemoveEmptyParagraphs
    }
}
