namespace OfficeIMO.Excel {
    /// <summary>
    /// Mutable worksheet view options such as zoom, gridlines, direction, and view mode.
    /// </summary>
    public sealed class ExcelWorksheetViewOptions {
        /// <summary>Gets or sets whether gridlines are visible.</summary>
        public bool? ShowGridlines { get; set; }

        /// <summary>Gets or sets whether the worksheet is shown right-to-left.</summary>
        public bool? RightToLeft { get; set; }

        /// <summary>Gets or sets the active zoom percentage. Excel supports values from 10 to 400.</summary>
        public uint? ZoomScale { get; set; }

        /// <summary>Gets or sets the normal-view zoom percentage. Excel supports values from 10 to 400.</summary>
        public uint? ZoomScaleNormal { get; set; }

        /// <summary>Gets or sets the worksheet view mode.</summary>
        public ExcelWorksheetViewKind? View { get; set; }
    }
}
