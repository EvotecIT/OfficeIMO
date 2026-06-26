namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only AxcExt date-axis extension metadata from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartAxisExtension {
        /// <summary>
        /// Creates date-axis extension metadata.
        /// </summary>
        public LegacyXlsChartAxisExtension(
            ushort minimumDate,
            ushort maximumDate,
            ushort majorInterval,
            ushort majorUnit,
            ushort minorInterval,
            ushort minorUnit,
            ushort baseUnit,
            ushort crossingDate,
            byte flags,
            byte reserved) {
            MinimumDate = minimumDate;
            MaximumDate = maximumDate;
            MajorInterval = majorInterval;
            MajorUnit = majorUnit;
            MinorInterval = minorInterval;
            MinorUnit = minorUnit;
            BaseUnit = baseUnit;
            CrossingDate = crossingDate;
            Flags = flags;
            Reserved = reserved;
        }

        /// <summary>Gets the raw minimum date value.</summary>
        public ushort MinimumDate { get; }

        /// <summary>Gets the raw maximum date value.</summary>
        public ushort MaximumDate { get; }

        /// <summary>Gets the raw major tick interval.</summary>
        public ushort MajorInterval { get; }

        /// <summary>Gets the raw major tick date unit.</summary>
        public ushort MajorUnit { get; }

        /// <summary>Gets the decoded major tick date unit name.</summary>
        public string MajorUnitName => GetDateUnitName(MajorUnit);

        /// <summary>Gets the raw minor tick interval.</summary>
        public ushort MinorInterval { get; }

        /// <summary>Gets the raw minor tick date unit.</summary>
        public ushort MinorUnit { get; }

        /// <summary>Gets the decoded minor tick date unit name.</summary>
        public string MinorUnitName => GetDateUnitName(MinorUnit);

        /// <summary>Gets the raw base date unit.</summary>
        public ushort BaseUnit { get; }

        /// <summary>Gets the decoded base date unit name.</summary>
        public string BaseUnitName => GetDateUnitName(BaseUnit);

        /// <summary>Gets the raw date at which the value axis crosses this axis.</summary>
        public ushort CrossingDate { get; }

        /// <summary>Gets the raw AxcExt flag byte.</summary>
        public byte Flags { get; }

        /// <summary>Gets whether the minimum date is automatic.</summary>
        public bool AutoMinimum => (Flags & 0x01) != 0;

        /// <summary>Gets whether the maximum date is automatic.</summary>
        public bool AutoMaximum => (Flags & 0x02) != 0;

        /// <summary>Gets whether the major interval is automatic.</summary>
        public bool AutoMajor => (Flags & 0x04) != 0;

        /// <summary>Gets whether the minor interval is automatic.</summary>
        public bool AutoMinor => (Flags & 0x08) != 0;

        /// <summary>Gets whether the axis is a date axis.</summary>
        public bool DateAxis => (Flags & 0x10) != 0;

        /// <summary>Gets whether the base date unit is automatic.</summary>
        public bool AutoBase => (Flags & 0x20) != 0;

        /// <summary>Gets whether the crossing date is automatic.</summary>
        public bool AutoCrossing => (Flags & 0x40) != 0;

        /// <summary>Gets whether Excel can automatically choose date-axis behavior.</summary>
        public bool AutoDateAxis => (Flags & 0x80) != 0;

        /// <summary>Gets the reserved byte that should be zero.</summary>
        public byte Reserved { get; }

        /// <summary>Gets whether the reserved byte is zero.</summary>
        public bool HasZeroReservedByte => Reserved == 0;

        private static string GetDateUnitName(ushort value) {
            switch (value) {
                case 0x0000:
                    return "Days";
                case 0x0001:
                    return "Months";
                case 0x0002:
                    return "Years";
                default:
                    return $"Unknown:0x{value:X4}";
            }
        }
    }
}
