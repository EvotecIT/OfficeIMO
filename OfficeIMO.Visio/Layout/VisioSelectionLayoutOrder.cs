namespace OfficeIMO.Visio {
    /// <summary>
    /// Ordering strategy used when relaying out an existing shape selection.
    /// </summary>
    public enum VisioSelectionLayoutOrder {
        /// <summary>Keep the order in which the selection was materialized.</summary>
        SelectionOrder = 0,

        /// <summary>Sort by page position from top-left to bottom-right.</summary>
        TopLeftToBottomRight = 1,

        /// <summary>Sort by page position from left-top to right-bottom.</summary>
        LeftTopToRightBottom = 2
    }
}
