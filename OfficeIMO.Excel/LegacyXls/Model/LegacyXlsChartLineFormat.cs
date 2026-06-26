namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Decoded metadata from a legacy XLS chart LineFormat record.
    /// </summary>
    public sealed class LegacyXlsChartLineFormat {
        /// <summary>
        /// Creates decoded chart line-format metadata.
        /// </summary>
        public LegacyXlsChartLineFormat(
            string rgbHex,
            ushort style,
            string styleName,
            short weight,
            string weightName,
            bool automatic,
            bool axisVisible,
            bool automaticColor,
            ushort colorIndex) {
            RgbHex = rgbHex ?? throw new ArgumentNullException(nameof(rgbHex));
            Style = style;
            StyleName = styleName ?? throw new ArgumentNullException(nameof(styleName));
            Weight = weight;
            WeightName = weightName ?? throw new ArgumentNullException(nameof(weightName));
            Automatic = automatic;
            AxisVisible = axisVisible;
            AutomaticColor = automaticColor;
            ColorIndex = colorIndex;
        }

        /// <summary>Gets the decoded RGB color in hexadecimal form.</summary>
        public string RgbHex { get; }

        /// <summary>Gets the raw line style code.</summary>
        public ushort Style { get; }

        /// <summary>Gets the decoded line style name.</summary>
        public string StyleName { get; }

        /// <summary>Gets the raw line weight code.</summary>
        public short Weight { get; }

        /// <summary>Gets the decoded line weight name.</summary>
        public string WeightName { get; }

        /// <summary>Gets whether the line uses automatic formatting.</summary>
        public bool Automatic { get; }

        /// <summary>Gets whether the axis line is displayed when the record applies to an axis line.</summary>
        public bool AxisVisible { get; }

        /// <summary>Gets whether the color index uses the automatic chart color.</summary>
        public bool AutomaticColor { get; }

        /// <summary>Gets the chart color index.</summary>
        public ushort ColorIndex { get; }
    }
}
