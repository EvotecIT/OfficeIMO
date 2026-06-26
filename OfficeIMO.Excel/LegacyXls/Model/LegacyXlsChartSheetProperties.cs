namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded ShtProps chart properties preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartSheetProperties {
        internal LegacyXlsChartSheetProperties(ushort flags, byte emptyCellPlottingMode) {
            Flags = flags;
            EmptyCellPlottingMode = emptyCellPlottingMode;
            EmptyCellPlottingModeName = GetEmptyCellPlottingModeName(emptyCellPlottingMode);
        }

        /// <summary>Gets the raw ShtProps flag bitfield.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether series are automatically allocated for the chart.</summary>
        public bool AutomaticallyAllocateSeries => (Flags & 0x0001) != 0;

        /// <summary>Gets whether only visible cells are plotted.</summary>
        public bool PlotVisibleCellsOnly => (Flags & 0x0002) != 0;

        /// <summary>Gets whether the chart is not sized with the window.</summary>
        public bool DoNotSizeWithWindow => (Flags & 0x0004) != 0;

        /// <summary>Gets whether manual plot-area dimensions are present.</summary>
        public bool ManualPlotArea => (Flags & 0x0008) != 0;

        /// <summary>Gets whether the default plot-area dimension is used.</summary>
        public bool AlwaysAutoPlotArea => (Flags & 0x0010) != 0;

        /// <summary>Gets the raw empty-cell plotting mode.</summary>
        public byte EmptyCellPlottingMode { get; }

        /// <summary>Gets whether the empty-cell plotting mode is one of the values defined by MS-XLS.</summary>
        public bool HasKnownEmptyCellPlottingMode => EmptyCellPlottingMode <= 0x02;

        /// <summary>Gets the decoded empty-cell plotting mode name.</summary>
        public string EmptyCellPlottingModeName { get; }

        private static string GetEmptyCellPlottingModeName(byte value) {
            switch (value) {
                case 0x00:
                    return "NotPlotted";
                case 0x01:
                    return "Zero";
                case 0x02:
                    return "Interpolated";
                default:
                    return $"Unknown:0x{value:X2}";
            }
        }
    }
}
