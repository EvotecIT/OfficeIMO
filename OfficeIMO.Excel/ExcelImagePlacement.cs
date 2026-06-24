namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes how a two-cell anchored worksheet image should react when rows or columns move or resize.
    /// </summary>
    public enum ExcelImagePlacement {
        /// <summary>The image moves and resizes with the cells between its start and end markers.</summary>
        MoveAndSize,
        /// <summary>The image moves with cells but keeps its current size.</summary>
        MoveOnly,
        /// <summary>The image keeps its absolute visual position relative to the worksheet drawing canvas.</summary>
        FreeFloating
    }
}
