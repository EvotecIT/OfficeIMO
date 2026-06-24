namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only CrtLayout12A plot-area layout metadata.
    /// </summary>
    public sealed class LegacyXlsChartPlotAreaLayout12 {
        internal LegacyXlsChartPlotAreaLayout12(
            uint checksum,
            bool targetsInnerPlotArea,
            short upperLeftX,
            short upperLeftY,
            short widthSprc,
            short heightSprc,
            ushort xMode,
            ushort yMode,
            ushort widthMode,
            ushort heightMode,
            double x,
            double y,
            double width,
            double height) {
            Checksum = checksum;
            TargetsInnerPlotArea = targetsInnerPlotArea;
            UpperLeftX = upperLeftX;
            UpperLeftY = upperLeftY;
            WidthSprc = widthSprc;
            HeightSprc = heightSprc;
            XMode = xMode;
            YMode = yMode;
            WidthMode = widthMode;
            HeightMode = heightMode;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>Gets the CrtLayout12A checksum value.</summary>
        public uint Checksum { get; }

        /// <summary>Gets whether the layout target is the inner plot area.</summary>
        public bool TargetsInnerPlotArea { get; }

        /// <summary>Gets the decoded plot-area target name.</summary>
        public string TargetName => TargetsInnerPlotArea ? "InnerPlotArea" : "OuterPlotArea";

        /// <summary>Gets the horizontal offset of the plot area's upper-left corner in SPRC units.</summary>
        public short UpperLeftX { get; }

        /// <summary>Gets the vertical offset of the plot area's upper-left corner in SPRC units.</summary>
        public short UpperLeftY { get; }

        /// <summary>Gets the plot-area width in SPRC units.</summary>
        public short WidthSprc { get; }

        /// <summary>Gets the plot-area height in SPRC units.</summary>
        public short HeightSprc { get; }

        /// <summary>Gets the raw X layout mode.</summary>
        public ushort XMode { get; }

        /// <summary>Gets the decoded X layout mode name.</summary>
        public string XModeName => LegacyXlsChartLayoutModeName.GetName(XMode);

        /// <summary>Gets the raw Y layout mode.</summary>
        public ushort YMode { get; }

        /// <summary>Gets the decoded Y layout mode name.</summary>
        public string YModeName => LegacyXlsChartLayoutModeName.GetName(YMode);

        /// <summary>Gets the raw width layout mode.</summary>
        public ushort WidthMode { get; }

        /// <summary>Gets the decoded width layout mode name.</summary>
        public string WidthModeName => LegacyXlsChartLayoutModeName.GetName(WidthMode);

        /// <summary>Gets the raw height layout mode.</summary>
        public ushort HeightMode { get; }

        /// <summary>Gets the decoded height layout mode name.</summary>
        public string HeightModeName => LegacyXlsChartLayoutModeName.GetName(HeightMode);

        /// <summary>Gets the X layout value.</summary>
        public double X { get; }

        /// <summary>Gets the Y layout value.</summary>
        public double Y { get; }

        /// <summary>Gets the width or lower-right X layout value.</summary>
        public double Width { get; }

        /// <summary>Gets the height or lower-right Y layout value.</summary>
        public double Height { get; }
    }
}
