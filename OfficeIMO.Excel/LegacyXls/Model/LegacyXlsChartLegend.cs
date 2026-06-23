namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart Legend record.
    /// </summary>
    public sealed class LegacyXlsChartLegend {
        internal LegacyXlsChartLegend(uint x, uint y, uint width, uint height, byte spacing, ushort flags) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Spacing = spacing;
            Flags = flags;
            AutoPosition = (flags & 0x0001) != 0;
            AutoPositionX = (flags & 0x0004) != 0;
            AutoPositionY = (flags & 0x0008) != 0;
            Vertical = (flags & 0x0010) != 0;
            WasDataTable = (flags & 0x0020) != 0;
        }

        /// <summary>Gets the legend X coordinate in SPRC units.</summary>
        public uint X { get; }

        /// <summary>Gets the legend Y coordinate in SPRC units.</summary>
        public uint Y { get; }

        /// <summary>Gets the legend width in SPRC units.</summary>
        public uint Width { get; }

        /// <summary>Gets the legend height in SPRC units.</summary>
        public uint Height { get; }

        /// <summary>Gets the raw spacing value between legend entries.</summary>
        public byte Spacing { get; }

        /// <summary>Gets the raw Legend record flag bitfield.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the legend position is automatic.</summary>
        public bool AutoPosition { get; }

        /// <summary>Gets whether the legend X position is automatic.</summary>
        public bool AutoPositionX { get; }

        /// <summary>Gets whether the legend Y position is automatic.</summary>
        public bool AutoPositionY { get; }

        /// <summary>Gets whether legend entries are laid out in one column.</summary>
        public bool Vertical { get; }

        /// <summary>Gets whether the legend is shown in a data table.</summary>
        public bool WasDataTable { get; }
    }
}
