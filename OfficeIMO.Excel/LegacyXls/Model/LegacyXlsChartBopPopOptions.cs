namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded BopPop chart group options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartBopPopOptions {
        internal LegacyXlsChartBopPopOptions(byte subtype, bool automaticSplit, ushort split, short splitPosition, short splitPercent, short secondaryPieSizePercent, short gapPercent, double splitValue, ushort flags) {
            Subtype = subtype;
            AutomaticSplit = automaticSplit;
            Split = split;
            SplitPosition = splitPosition;
            SplitPercent = splitPercent;
            SecondaryPieSizePercent = secondaryPieSizePercent;
            GapPercent = gapPercent;
            SplitValue = splitValue;
            Flags = flags;
        }

        /// <summary>Gets the raw bar-of-pie or pie-of-pie subtype.</summary>
        public byte Subtype { get; }

        /// <summary>Gets the decoded bar-of-pie or pie-of-pie subtype name.</summary>
        public string SubtypeName => Subtype switch {
            0x01 => "PieOfPie",
            0x02 => "BarOfPie",
            _ => $"Unknown:0x{Subtype:X2}"
        };

        /// <summary>Gets whether the subtype is one of the BIFF-defined values.</summary>
        public bool HasKnownSubtype => Subtype is 0x01 or 0x02;

        /// <summary>Gets whether Excel automatically determines the secondary bar/pie split.</summary>
        public bool AutomaticSplit { get; }

        /// <summary>Gets the raw split mode.</summary>
        public ushort Split { get; }

        /// <summary>Gets the decoded split mode name.</summary>
        public string SplitName => Split switch {
            0x0000 => "Position",
            0x0001 => "Value",
            0x0002 => "Percent",
            0x0003 => "Custom",
            _ => $"Unknown:0x{Split:X4}"
        };

        /// <summary>Gets whether the split mode is one of the BIFF-defined values.</summary>
        public bool HasKnownSplit => Split <= 0x0003;

        /// <summary>Gets the secondary bar/pie split position.</summary>
        public short SplitPosition { get; }

        /// <summary>Gets the percentage threshold for percent-based splits.</summary>
        public short SplitPercent { get; }

        /// <summary>Gets the secondary bar/pie size as a percentage of the primary pie size.</summary>
        public short SecondaryPieSizePercent { get; }

        /// <summary>Gets the gap between the primary pie and secondary bar/pie.</summary>
        public short GapPercent { get; }

        /// <summary>Gets the threshold value for value-based splits.</summary>
        public double SplitValue { get; }

        /// <summary>Gets the raw BopPop flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether one or more data points have shadows.</summary>
        public bool HasShadow => (Flags & 0x0001) != 0;

        /// <summary>Gets whether the reserved BopPop flag bits are zero.</summary>
        public bool HasZeroReservedBits => (Flags & 0xfffe) == 0;
    }
}
