namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Specifies how a target size should be derived from a selection.
    /// </summary>
    public enum PowerPointShapeSizeReference {
        /// <summary>Use the first shape's size.</summary>
        First,
        /// <summary>Use the smallest width/height in the selection.</summary>
        Smallest,
        /// <summary>Use the largest width/height in the selection.</summary>
        Largest,
        /// <summary>Use the average width/height in the selection.</summary>
        Average
    }
}
