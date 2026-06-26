namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded PlotGrowth chart font-scaling metadata.
    /// </summary>
    public sealed class LegacyXlsChartPlotGrowth {
        internal LegacyXlsChartPlotGrowth(
            short horizontalIntegral,
            ushort horizontalFractional,
            short verticalIntegral,
            ushort verticalFractional) {
            HorizontalIntegral = horizontalIntegral;
            HorizontalFractional = horizontalFractional;
            VerticalIntegral = verticalIntegral;
            VerticalFractional = verticalFractional;
        }

        /// <summary>Gets the integral part of the horizontal growth factor.</summary>
        public short HorizontalIntegral { get; }

        /// <summary>Gets the fractional part of the horizontal growth factor.</summary>
        public ushort HorizontalFractional { get; }

        /// <summary>Gets the horizontal growth factor in points.</summary>
        public double HorizontalGrowthPoints => HorizontalIntegral + HorizontalFractional / 65536d;

        /// <summary>Gets the integral part of the vertical growth factor.</summary>
        public short VerticalIntegral { get; }

        /// <summary>Gets the fractional part of the vertical growth factor.</summary>
        public ushort VerticalFractional { get; }

        /// <summary>Gets the vertical growth factor in points.</summary>
        public double VerticalGrowthPoints => VerticalIntegral + VerticalFractional / 65536d;
    }
}
