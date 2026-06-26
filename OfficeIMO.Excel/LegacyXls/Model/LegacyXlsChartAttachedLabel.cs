namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes data-label display options decoded from a legacy XLS chart AttachedLabel record.
    /// </summary>
    public sealed class LegacyXlsChartAttachedLabel {
        internal LegacyXlsChartAttachedLabel(ushort flags) {
            Flags = flags;
        }

        /// <summary>Gets the raw AttachedLabel flag field.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the data label displays the value.</summary>
        public bool ShowValue => (Flags & 0x0001) != 0;

        /// <summary>Gets whether the data label displays a percentage.</summary>
        public bool ShowPercent => (Flags & 0x0002) != 0;

        /// <summary>Gets whether the data label displays both category label and percentage.</summary>
        public bool ShowLabelAndPercent => (Flags & 0x0004) != 0;

        /// <summary>Gets whether the data label displays the category label.</summary>
        public bool ShowLabel => (Flags & 0x0010) != 0;

        /// <summary>Gets whether the data label displays bubble sizes.</summary>
        public bool ShowBubbleSizes => (Flags & 0x0020) != 0;

        /// <summary>Gets whether the data label displays the series name.</summary>
        public bool ShowSeriesName => (Flags & 0x0040) != 0;

        /// <summary>Gets decoded flag names for report grouping.</summary>
        public IReadOnlyList<string> FlagNames {
            get {
                var names = new List<string>();
                if (ShowValue) names.Add("ShowValue");
                if (ShowPercent) names.Add("ShowPercent");
                if (ShowLabelAndPercent) names.Add("ShowLabelAndPercent");
                if (ShowLabel) names.Add("ShowLabel");
                if (ShowBubbleSizes) names.Add("ShowBubbleSizes");
                if (ShowSeriesName) names.Add("ShowSeriesName");
                return names;
            }
        }
    }
}
