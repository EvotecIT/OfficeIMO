namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart Frame record.
    /// </summary>
    public sealed class LegacyXlsChartFrame {
        internal LegacyXlsChartFrame(ushort frameType, string frameTypeName, ushort flags, bool automaticSize, bool automaticPosition) {
            FrameType = frameType;
            FrameTypeName = frameTypeName ?? throw new ArgumentNullException(nameof(frameTypeName));
            Flags = flags;
            AutomaticSize = automaticSize;
            AutomaticPosition = automaticPosition;
        }

        /// <summary>Gets the raw frame type.</summary>
        public ushort FrameType { get; }

        /// <summary>Gets the decoded frame type name.</summary>
        public string FrameTypeName { get; }

        /// <summary>Gets the raw frame flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the frame size is automatically calculated.</summary>
        public bool AutomaticSize { get; }

        /// <summary>Gets whether the frame position is automatically calculated.</summary>
        public bool AutomaticPosition { get; }
    }
}
