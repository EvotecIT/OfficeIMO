namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Chart3DBarShape options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChart3DBarShapeOptions {
        internal LegacyXlsChart3DBarShapeOptions(byte riser, byte taper) {
            Riser = riser;
            Taper = taper;
        }

        /// <summary>Gets the raw base-shape value for the chart data points.</summary>
        public byte Riser { get; }

        /// <summary>Gets the decoded base-shape name for the chart data points.</summary>
        public string RiserName => Riser switch {
            0x00 => "Rectangle",
            0x01 => "Ellipse",
            _ => $"Unknown:0x{Riser:X2}"
        };

        /// <summary>Gets the raw tapering value for the chart data points.</summary>
        public byte Taper { get; }

        /// <summary>Gets the decoded tapering mode name for the chart data points.</summary>
        public string TaperName => Taper switch {
            0x00 => "None",
            0x01 => "Point",
            0x02 => "ProjectedPoint",
            _ => $"Unknown:0x{Taper:X2}"
        };
    }
}
