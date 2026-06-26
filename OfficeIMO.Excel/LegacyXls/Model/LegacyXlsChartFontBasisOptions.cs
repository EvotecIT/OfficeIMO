namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Fbi or Fbi2 chart font-scaling metadata preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartFontBasisOptions {
        internal LegacyXlsChartFontBasisOptions(ushort widthTwipsBasis, ushort heightTwipsBasis, ushort fontHeightTwips, ushort scaleBasis, ushort fontIndex) {
            WidthTwipsBasis = widthTwipsBasis;
            HeightTwipsBasis = heightTwipsBasis;
            FontHeightTwips = fontHeightTwips;
            ScaleBasis = scaleBasis;
            FontIndex = fontIndex;
        }

        /// <summary>Gets the chart width, in twips, when the scalable font was first applied.</summary>
        public ushort WidthTwipsBasis { get; }

        /// <summary>Gets the chart height, in twips, when the scalable font was first applied.</summary>
        public ushort HeightTwipsBasis { get; }

        /// <summary>Gets the default font height, in twips.</summary>
        public ushort FontHeightTwips { get; }

        /// <summary>Gets the raw scale basis value.</summary>
        public ushort ScaleBasis { get; }

        /// <summary>Gets the decoded scale basis name.</summary>
        public string ScaleBasisName => ScaleBasis switch {
            0x0000 => "ChartArea",
            0x0001 => "PlotArea",
            _ => $"Unknown:0x{ScaleBasis:X4}"
        };

        /// <summary>Gets whether the scale basis is one of the BIFF-defined values.</summary>
        public bool HasKnownScaleBasis => ScaleBasis is 0x0000 or 0x0001;

        /// <summary>Gets the referenced chart font index.</summary>
        public ushort FontIndex { get; }
    }
}
