namespace OfficeIMO.Excel {
    /// <summary>
    /// Read-only worksheet view metadata such as frozen panes, gridlines, zoom, and direction.
    /// </summary>
    public sealed class ExcelWorksheetViewInfo {
        /// <summary>Gets or sets whether the worksheet contains a pane definition.</summary>
        public bool HasPane { get; internal set; }

        /// <summary>Gets or sets the pane state, for example <c>frozen</c> or <c>split</c>.</summary>
        public string? PaneState { get; internal set; }

        /// <summary>Gets or sets the horizontal split value stored by Excel.</summary>
        public double? HorizontalSplit { get; internal set; }

        /// <summary>Gets or sets the vertical split value stored by Excel.</summary>
        public double? VerticalSplit { get; internal set; }

        /// <summary>Gets or sets the number of frozen rows, when the pane is frozen.</summary>
        public int FrozenRowCount { get; internal set; }

        /// <summary>Gets or sets the number of frozen columns, when the pane is frozen.</summary>
        public int FrozenColumnCount { get; internal set; }

        /// <summary>Gets or sets the top-left scrollable cell, when present.</summary>
        public string? TopLeftCell { get; internal set; }

        /// <summary>Gets or sets the active pane, when present.</summary>
        public string? ActivePane { get; internal set; }

        /// <summary>Gets or sets whether worksheet gridlines are visible.</summary>
        public bool ShowGridlines { get; internal set; } = true;

        /// <summary>Gets or sets whether the worksheet view is right-to-left.</summary>
        public bool RightToLeft { get; internal set; }

        /// <summary>Gets or sets the sheet view type, when present.</summary>
        public string? View { get; internal set; }

        /// <summary>Gets or sets the current zoom scale, when present.</summary>
        public uint? ZoomScale { get; internal set; }

        /// <summary>Gets or sets the normal-view zoom scale, when present.</summary>
        public uint? ZoomScaleNormal { get; internal set; }
    }
}
