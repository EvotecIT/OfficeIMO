namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Decoded metadata from a legacy XLS chart MarkerFormat record.
    /// </summary>
    public sealed class LegacyXlsChartMarkerFormat {
        /// <summary>
        /// Creates decoded chart marker-format metadata.
        /// </summary>
        public LegacyXlsChartMarkerFormat(
            string foregroundRgbHex,
            string backgroundRgbHex,
            ushort markerType,
            string markerTypeName,
            bool automatic,
            bool interiorHidden,
            bool borderHidden,
            ushort foregroundColorIndex,
            ushort backgroundColorIndex,
            uint sizeTwips) {
            ForegroundRgbHex = foregroundRgbHex ?? throw new ArgumentNullException(nameof(foregroundRgbHex));
            BackgroundRgbHex = backgroundRgbHex ?? throw new ArgumentNullException(nameof(backgroundRgbHex));
            MarkerType = markerType;
            MarkerTypeName = markerTypeName ?? throw new ArgumentNullException(nameof(markerTypeName));
            Automatic = automatic;
            InteriorHidden = interiorHidden;
            BorderHidden = borderHidden;
            ForegroundColorIndex = foregroundColorIndex;
            BackgroundColorIndex = backgroundColorIndex;
            SizeTwips = sizeTwips;
        }

        /// <summary>Gets the decoded marker border RGB color in hexadecimal form.</summary>
        public string ForegroundRgbHex { get; }

        /// <summary>Gets the decoded marker interior RGB color in hexadecimal form.</summary>
        public string BackgroundRgbHex { get; }

        /// <summary>Gets the raw marker type code.</summary>
        public ushort MarkerType { get; }

        /// <summary>Gets the decoded marker type name.</summary>
        public string MarkerTypeName { get; }

        /// <summary>Gets whether marker formatting is automatically selected.</summary>
        public bool Automatic { get; }

        /// <summary>Gets whether the marker interior is hidden.</summary>
        public bool InteriorHidden { get; }

        /// <summary>Gets whether the marker border is hidden.</summary>
        public bool BorderHidden { get; }

        /// <summary>Gets the chart foreground color index.</summary>
        public ushort ForegroundColorIndex { get; }

        /// <summary>Gets the chart background color index.</summary>
        public ushort BackgroundColorIndex { get; }

        /// <summary>Gets the marker size in twips.</summary>
        public uint SizeTwips { get; }
    }
}
