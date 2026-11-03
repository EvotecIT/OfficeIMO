using System;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Defines chart style/color presets for Excel charts.
    /// </summary>
    public sealed class ExcelChartStylePreset {
        /// <summary>
        /// Default chart preset (style 251, color 10).
        /// </summary>
        public static ExcelChartStylePreset Default { get; } = new ExcelChartStylePreset();

        /// <summary>
        /// Creates a preset that uses built-in chart style/color resources.
        /// </summary>
        public ExcelChartStylePreset(int styleId = 251, int colorStyleId = 10) {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(styleId);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(colorStyleId);
            StyleId = styleId;
            ColorStyleId = colorStyleId;
        }

        /// <summary>
        /// Creates a preset from custom chart style and color XML.
        /// </summary>
        public ExcelChartStylePreset(string styleXml, string colorStyleXml, int styleId = 251, int colorStyleId = 10)
            : this(styleId, colorStyleId) {
            ArgumentNullException.ThrowIfNull(styleXml);
            ArgumentNullException.ThrowIfNull(colorStyleXml);
            StyleXmlBytes = Encoding.UTF8.GetBytes(styleXml);
            ColorXmlBytes = Encoding.UTF8.GetBytes(colorStyleXml);
        }

        /// <summary>
        /// Creates a preset from custom chart style and color XML bytes.
        /// </summary>
        public ExcelChartStylePreset(byte[] styleXmlBytes, byte[] colorStyleXmlBytes, int styleId = 251, int colorStyleId = 10)
            : this(styleId, colorStyleId) {
            ArgumentNullException.ThrowIfNull(styleXmlBytes);
            ArgumentNullException.ThrowIfNull(colorStyleXmlBytes);
            StyleXmlBytes = (byte[])styleXmlBytes.Clone();
            ColorXmlBytes = (byte[])colorStyleXmlBytes.Clone();
        }

        /// <summary>
        /// Gets the chart style id.
        /// </summary>
        public int StyleId { get; }

        /// <summary>
        /// Gets the chart color style id.
        /// </summary>
        public int ColorStyleId { get; }

        internal byte[]? StyleXmlBytes { get; }
        internal byte[]? ColorXmlBytes { get; }
    }
}
