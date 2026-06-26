namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only CatLab axis-label metadata from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartCategoryLabelOptions {
        internal LegacyXlsChartCategoryLabelOptions(ushort offsetPercentage, ushort alignment, bool useAutomaticLabelCount) {
            OffsetPercentage = offsetPercentage;
            Alignment = alignment;
            UseAutomaticLabelCount = useAutomaticLabelCount;
        }

        /// <summary>Gets the axis-label offset as a percentage of the default distance.</summary>
        public ushort OffsetPercentage { get; }

        /// <summary>Gets the raw axis-label alignment value.</summary>
        public ushort Alignment { get; }

        /// <summary>Gets the decoded axis-label alignment name.</summary>
        public string AlignmentName => Alignment switch {
            0x0001 => "TopOrReadingOrderStart",
            0x0002 => "Center",
            0x0003 => "BottomOrReadingOrderEnd",
            _ => $"Unknown:0x{Alignment:X4}"
        };

        /// <summary>Gets whether the category label count is automatically calculated.</summary>
        public bool UseAutomaticLabelCount { get; }
    }
}
