namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only CrtLayout12 chart layout metadata.
    /// </summary>
    public sealed class LegacyXlsChartLayout12 {
        internal LegacyXlsChartLayout12(
            uint checksum,
            byte automaticLayoutType,
            ushort xMode,
            ushort yMode,
            ushort widthMode,
            ushort heightMode,
            double x,
            double y,
            double width,
            double height) {
            Checksum = checksum;
            AutomaticLayoutType = automaticLayoutType;
            XMode = xMode;
            YMode = yMode;
            WidthMode = widthMode;
            HeightMode = heightMode;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>Gets the CrtLayout12 checksum value.</summary>
        public uint Checksum { get; }

        /// <summary>Gets the automatic legend layout type.</summary>
        public byte AutomaticLayoutType { get; }

        /// <summary>Gets the decoded automatic legend layout type name.</summary>
        public string AutomaticLayoutTypeName => AutomaticLayoutType switch {
            0x00 => "Bottom",
            0x01 => "TopRight",
            0x02 => "Top",
            0x03 => "Right",
            0x04 => "Left",
            _ => $"Unknown:0x{AutomaticLayoutType:X2}"
        };

        /// <summary>Gets the raw X layout mode.</summary>
        public ushort XMode { get; }

        /// <summary>Gets the decoded X layout mode name.</summary>
        public string XModeName => GetModeName(XMode);

        /// <summary>Gets the raw Y layout mode.</summary>
        public ushort YMode { get; }

        /// <summary>Gets the decoded Y layout mode name.</summary>
        public string YModeName => GetModeName(YMode);

        /// <summary>Gets the raw width layout mode.</summary>
        public ushort WidthMode { get; }

        /// <summary>Gets the decoded width layout mode name.</summary>
        public string WidthModeName => GetModeName(WidthMode);

        /// <summary>Gets the raw height layout mode.</summary>
        public ushort HeightMode { get; }

        /// <summary>Gets the decoded height layout mode name.</summary>
        public string HeightModeName => GetModeName(HeightMode);

        /// <summary>Gets the X layout value.</summary>
        public double X { get; }

        /// <summary>Gets the Y layout value.</summary>
        public double Y { get; }

        /// <summary>Gets the width or lower-right X layout value.</summary>
        public double Width { get; }

        /// <summary>Gets the height or lower-right Y layout value.</summary>
        public double Height { get; }

        private static string GetModeName(ushort mode) {
            return mode switch {
                0x0000 => "Automatic",
                0x0001 => "Factor",
                0x0002 => "Edge",
                _ => $"Unknown:0x{mode:X4}"
            };
        }
    }
}
