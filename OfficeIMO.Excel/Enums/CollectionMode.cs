namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how collections are handled when flattening objects.
    /// </summary>
    public enum CollectionMode {
        /// <summary>
        /// Join collection items into a single cell using a separator.
        /// </summary>
        JoinWith,

        /// <summary>
        /// Expand each collection item into its own row.
        /// </summary>
        ExpandRows
    }
}

