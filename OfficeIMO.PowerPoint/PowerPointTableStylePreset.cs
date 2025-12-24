namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Describes a table style preset that toggles built-in table flags.
    /// </summary>
    public readonly struct PowerPointTableStylePreset {
        /// <summary>
        ///     Creates a table style preset with optional style flags.
        /// </summary>
        public PowerPointTableStylePreset(string? styleId = null, bool? firstRow = null, bool? lastRow = null,
            bool? firstColumn = null, bool? lastColumn = null, bool? bandedRows = null, bool? bandedColumns = null) {
            StyleId = styleId;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            BandedRows = bandedRows;
            BandedColumns = bandedColumns;
        }

        /// <summary>
        ///     Optional style ID to apply.
        /// </summary>
        public string? StyleId { get; }

        /// <summary>
        ///     Indicates whether the first row styling should be enabled.
        /// </summary>
        public bool? FirstRow { get; }

        /// <summary>
        ///     Indicates whether the last row styling should be enabled.
        /// </summary>
        public bool? LastRow { get; }

        /// <summary>
        ///     Indicates whether the first column styling should be enabled.
        /// </summary>
        public bool? FirstColumn { get; }

        /// <summary>
        ///     Indicates whether the last column styling should be enabled.
        /// </summary>
        public bool? LastColumn { get; }

        /// <summary>
        ///     Indicates whether banded rows styling should be enabled.
        /// </summary>
        public bool? BandedRows { get; }

        /// <summary>
        ///     Indicates whether banded columns styling should be enabled.
        /// </summary>
        public bool? BandedColumns { get; }

        /// <summary>
        ///     Default table preset (first row + banded rows).
        /// </summary>
        public static PowerPointTableStylePreset Default =>
            new PowerPointTableStylePreset(firstRow: true, bandedRows: true);

        /// <summary>
        ///     Plain table preset without header or banding.
        /// </summary>
        public static PowerPointTableStylePreset Plain =>
            new PowerPointTableStylePreset(firstRow: false, bandedRows: false);

        /// <summary>
        ///     Header-only preset.
        /// </summary>
        public static PowerPointTableStylePreset HeaderOnly =>
            new PowerPointTableStylePreset(firstRow: true, bandedRows: false);

        /// <summary>
        ///     Banded rows preset.
        /// </summary>
        public static PowerPointTableStylePreset BandedRowsOnly =>
            new PowerPointTableStylePreset(bandedRows: true);

        /// <summary>
        ///     Banded columns preset.
        /// </summary>
        public static PowerPointTableStylePreset BandedColumnsOnly =>
            new PowerPointTableStylePreset(bandedColumns: true);
    }
}
