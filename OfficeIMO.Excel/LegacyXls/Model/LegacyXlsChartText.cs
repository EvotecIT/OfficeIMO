namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart Text record.
    /// </summary>
    public sealed class LegacyXlsChartText {
        internal LegacyXlsChartText(
            byte horizontalAlignment,
            string horizontalAlignmentName,
            byte verticalAlignment,
            string verticalAlignmentName,
            ushort backgroundMode,
            string backgroundModeName,
            string rgbHex,
            int x,
            int y,
            int width,
            int height,
            ushort flags,
            IReadOnlyList<string> flagNames,
            ushort colorIndex,
            byte dataLabelPosition,
            string dataLabelPositionName,
            byte readingOrder,
            string readingOrderName,
            ushort rotation) {
            HorizontalAlignment = horizontalAlignment;
            HorizontalAlignmentName = horizontalAlignmentName ?? throw new ArgumentNullException(nameof(horizontalAlignmentName));
            VerticalAlignment = verticalAlignment;
            VerticalAlignmentName = verticalAlignmentName ?? throw new ArgumentNullException(nameof(verticalAlignmentName));
            BackgroundMode = backgroundMode;
            BackgroundModeName = backgroundModeName ?? throw new ArgumentNullException(nameof(backgroundModeName));
            RgbHex = rgbHex ?? throw new ArgumentNullException(nameof(rgbHex));
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Flags = flags;
            FlagNames = flagNames ?? throw new ArgumentNullException(nameof(flagNames));
            ColorIndex = colorIndex;
            DataLabelPosition = dataLabelPosition;
            DataLabelPositionName = dataLabelPositionName ?? throw new ArgumentNullException(nameof(dataLabelPositionName));
            ReadingOrder = readingOrder;
            ReadingOrderName = readingOrderName ?? throw new ArgumentNullException(nameof(readingOrderName));
            Rotation = rotation;
        }

        /// <summary>Gets the raw horizontal alignment identifier.</summary>
        public byte HorizontalAlignment { get; }

        /// <summary>Gets the decoded horizontal alignment name.</summary>
        public string HorizontalAlignmentName { get; }

        /// <summary>Gets the raw vertical alignment identifier.</summary>
        public byte VerticalAlignment { get; }

        /// <summary>Gets the decoded vertical alignment name.</summary>
        public string VerticalAlignmentName { get; }

        /// <summary>Gets the raw text background mode.</summary>
        public ushort BackgroundMode { get; }

        /// <summary>Gets the decoded text background mode name.</summary>
        public string BackgroundModeName { get; }

        /// <summary>Gets the text foreground color as RGB hex.</summary>
        public string RgbHex { get; }

        /// <summary>Gets the horizontal text position in SPRC units.</summary>
        public int X { get; }

        /// <summary>Gets the vertical text position in SPRC units.</summary>
        public int Y { get; }

        /// <summary>Gets the text width in SPRC units.</summary>
        public int Width { get; }

        /// <summary>Gets the text height in SPRC units.</summary>
        public int Height { get; }

        /// <summary>Gets the raw Text record flag bitfield.</summary>
        public ushort Flags { get; }

        /// <summary>Gets decoded Text record flag names.</summary>
        public IReadOnlyList<string> FlagNames { get; }

        /// <summary>Gets the indexed text color.</summary>
        public ushort ColorIndex { get; }

        /// <summary>Gets the raw data-label position identifier.</summary>
        public byte DataLabelPosition { get; }

        /// <summary>Gets the decoded data-label position name.</summary>
        public string DataLabelPositionName { get; }

        /// <summary>Gets the raw reading-order identifier.</summary>
        public byte ReadingOrder { get; }

        /// <summary>Gets the decoded reading-order name.</summary>
        public string ReadingOrderName { get; }

        /// <summary>Gets the text rotation value.</summary>
        public ushort Rotation { get; }
    }
}
