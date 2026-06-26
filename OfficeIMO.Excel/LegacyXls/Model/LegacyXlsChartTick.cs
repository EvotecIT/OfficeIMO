namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart Tick record.
    /// </summary>
    public sealed class LegacyXlsChartTick {
        internal LegacyXlsChartTick(
            byte majorTickLocation,
            string majorTickLocationName,
            byte minorTickLocation,
            string minorTickLocationName,
            byte labelLocation,
            string labelLocationName,
            byte backgroundMode,
            string backgroundModeName,
            string rgbHex,
            ushort flags,
            byte rotationMode,
            string rotationModeName,
            bool autoColor,
            bool autoBackground,
            bool autoRotation,
            byte readingOrder,
            string readingOrderName,
            ushort colorIndex,
            ushort rotation) {
            MajorTickLocation = majorTickLocation;
            MajorTickLocationName = majorTickLocationName ?? throw new ArgumentNullException(nameof(majorTickLocationName));
            MinorTickLocation = minorTickLocation;
            MinorTickLocationName = minorTickLocationName ?? throw new ArgumentNullException(nameof(minorTickLocationName));
            LabelLocation = labelLocation;
            LabelLocationName = labelLocationName ?? throw new ArgumentNullException(nameof(labelLocationName));
            BackgroundMode = backgroundMode;
            BackgroundModeName = backgroundModeName ?? throw new ArgumentNullException(nameof(backgroundModeName));
            RgbHex = rgbHex ?? throw new ArgumentNullException(nameof(rgbHex));
            Flags = flags;
            RotationMode = rotationMode;
            RotationModeName = rotationModeName ?? throw new ArgumentNullException(nameof(rotationModeName));
            AutoColor = autoColor;
            AutoBackground = autoBackground;
            AutoRotation = autoRotation;
            ReadingOrder = readingOrder;
            ReadingOrderName = readingOrderName ?? throw new ArgumentNullException(nameof(readingOrderName));
            ColorIndex = colorIndex;
            Rotation = rotation;
        }

        /// <summary>Gets the raw major tick location identifier.</summary>
        public byte MajorTickLocation { get; }

        /// <summary>Gets the decoded major tick location name.</summary>
        public string MajorTickLocationName { get; }

        /// <summary>Gets the raw minor tick location identifier.</summary>
        public byte MinorTickLocation { get; }

        /// <summary>Gets the decoded minor tick location name.</summary>
        public string MinorTickLocationName { get; }

        /// <summary>Gets the raw axis-label location identifier.</summary>
        public byte LabelLocation { get; }

        /// <summary>Gets the decoded axis-label location name.</summary>
        public string LabelLocationName { get; }

        /// <summary>Gets the raw axis-label background mode.</summary>
        public byte BackgroundMode { get; }

        /// <summary>Gets the decoded axis-label background mode name.</summary>
        public string BackgroundModeName { get; }

        /// <summary>Gets the axis-label foreground color as RGB hex.</summary>
        public string RgbHex { get; }

        /// <summary>Gets the raw Tick record flag bitfield.</summary>
        public ushort Flags { get; }

        /// <summary>Gets the raw axis-label rotation mode.</summary>
        public byte RotationMode { get; }

        /// <summary>Gets the decoded axis-label rotation mode name.</summary>
        public string RotationModeName { get; }

        /// <summary>Gets whether the axis-label foreground color is automatic.</summary>
        public bool AutoColor { get; }

        /// <summary>Gets whether the axis-label background is automatic.</summary>
        public bool AutoBackground { get; }

        /// <summary>Gets whether axis-label rotation is automatic.</summary>
        public bool AutoRotation { get; }

        /// <summary>Gets the raw reading-order identifier.</summary>
        public byte ReadingOrder { get; }

        /// <summary>Gets the decoded reading-order name.</summary>
        public string ReadingOrderName { get; }

        /// <summary>Gets the indexed text color.</summary>
        public ushort ColorIndex { get; }

        /// <summary>Gets the axis-label text rotation value.</summary>
        public ushort Rotation { get; }
    }
}
