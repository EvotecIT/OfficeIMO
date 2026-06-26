namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Decoded metadata from a legacy XLS chart AreaFormat record.
    /// </summary>
    public sealed class LegacyXlsChartAreaFormat {
        /// <summary>
        /// Creates decoded chart area-format metadata.
        /// </summary>
        public LegacyXlsChartAreaFormat(
            string foregroundRgbHex,
            string backgroundRgbHex,
            ushort pattern,
            string patternName,
            bool automatic,
            bool invertNegative,
            ushort foregroundColorIndex,
            ushort backgroundColorIndex) {
            ForegroundRgbHex = foregroundRgbHex ?? throw new ArgumentNullException(nameof(foregroundRgbHex));
            BackgroundRgbHex = backgroundRgbHex ?? throw new ArgumentNullException(nameof(backgroundRgbHex));
            Pattern = pattern;
            PatternName = patternName ?? throw new ArgumentNullException(nameof(patternName));
            Automatic = automatic;
            InvertNegative = invertNegative;
            ForegroundColorIndex = foregroundColorIndex;
            BackgroundColorIndex = backgroundColorIndex;
        }

        /// <summary>Gets the decoded foreground RGB color in hexadecimal form.</summary>
        public string ForegroundRgbHex { get; }

        /// <summary>Gets the decoded background RGB color in hexadecimal form.</summary>
        public string BackgroundRgbHex { get; }

        /// <summary>Gets the raw fill pattern code.</summary>
        public ushort Pattern { get; }

        /// <summary>Gets the decoded fill pattern name.</summary>
        public string PatternName { get; }

        /// <summary>Gets whether fill colors are automatically selected.</summary>
        public bool Automatic { get; }

        /// <summary>Gets whether foreground and background colors are swapped for negative values.</summary>
        public bool InvertNegative { get; }

        /// <summary>Gets the chart foreground color index.</summary>
        public ushort ForegroundColorIndex { get; }

        /// <summary>Gets the chart background color index.</summary>
        public ushort BackgroundColorIndex { get; }
    }
}
