namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes properties decoded from a legacy XLS chart SerFmt record.
    /// </summary>
    public sealed class LegacyXlsChartSeriesFormat {
        internal LegacyXlsChartSeriesFormat(ushort flags) {
            Flags = flags;
        }

        /// <summary>Gets the raw SerFmt flag field.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether series lines use a smooth-line effect.</summary>
        public bool SmoothLine => (Flags & 0x0001) != 0;

        /// <summary>Gets whether bubble chart data points use a 3-D effect.</summary>
        public bool ThreeDimensionalBubbles => (Flags & 0x0002) != 0;

        /// <summary>Gets whether data markers use a shadow.</summary>
        public bool Shadow => (Flags & 0x0004) != 0;

        /// <summary>Gets the reserved SerFmt bits.</summary>
        public ushort Reserved => (ushort)(Flags & 0xfff8);

        /// <summary>Gets whether all reserved bits are zero.</summary>
        public bool HasZeroReservedBits => Reserved == 0;

        /// <summary>Gets decoded flag names for report grouping.</summary>
        public IReadOnlyList<string> FlagNames {
            get {
                var names = new List<string>();
                if (SmoothLine) names.Add("SmoothLine");
                if (ThreeDimensionalBubbles) names.Add("ThreeDimensionalBubbles");
                if (Shadow) names.Add("Shadow");
                return names;
            }
        }
    }
}
