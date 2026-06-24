namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Scatter chart group options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartScatterOptions {
        internal LegacyXlsChartScatterOptions(ushort bubbleSizeRatio, ushort bubbleSizeRepresentation, ushort flags) {
            BubbleSizeRatio = bubbleSizeRatio;
            BubbleSizeRepresentation = bubbleSizeRepresentation;
            Flags = flags;
        }

        /// <summary>Gets the bubble-size scale ratio as a percentage of the default bubble size.</summary>
        public ushort BubbleSizeRatio { get; }

        /// <summary>Gets whether bubble size is interpreted as area or width.</summary>
        public ushort BubbleSizeRepresentation { get; }

        /// <summary>Gets the decoded bubble-size representation name.</summary>
        public string BubbleSizeRepresentationName => BubbleSizeRepresentation switch {
            0x0001 => "Area",
            0x0002 => "Width",
            _ => $"Unknown:0x{BubbleSizeRepresentation:X4}"
        };

        /// <summary>Gets whether the bubble-size representation is one of the BIFF-defined values.</summary>
        public bool HasKnownBubbleSizeRepresentation => BubbleSizeRepresentation is 0x0001 or 0x0002;

        /// <summary>Gets whether the bubble-size scale ratio is in the BIFF-defined zero-to-300 percent range.</summary>
        public bool HasValidBubbleSizeRatio => BubbleSizeRatio <= 300;

        /// <summary>Gets the raw Scatter option flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the scatter chart group is displayed as bubbles.</summary>
        public bool IsBubbleChart => (Flags & 0x0001) != 0;

        /// <summary>Gets whether negative bubble-size values are displayed.</summary>
        public bool ShowNegativeBubbles => (Flags & 0x0002) != 0;

        /// <summary>Gets whether data points in the chart group have shadows.</summary>
        public bool HasShadow => (Flags & 0x0004) != 0;
    }
}
