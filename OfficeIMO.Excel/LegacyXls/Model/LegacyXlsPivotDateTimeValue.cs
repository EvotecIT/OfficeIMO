using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an SXDtr date/time value attached to legacy XLS PivotCache grouping metadata.
    /// </summary>
    public sealed class LegacyXlsPivotDateTimeValue {
        /// <summary>Creates a PivotCache date/time value.</summary>
        public LegacyXlsPivotDateTimeValue(ushort year, ushort month, byte day, byte hour, byte minute, byte second) {
            Year = year;
            Month = month;
            Day = day;
            Hour = hour;
            Minute = minute;
            Second = second;
        }

        /// <summary>Gets the year component.</summary>
        public ushort Year { get; }

        /// <summary>Gets the month component.</summary>
        public ushort Month { get; }

        /// <summary>Gets the day-of-month component, or zero for time-only cache values.</summary>
        public byte Day { get; }

        /// <summary>Gets the hour component.</summary>
        public byte Hour { get; }

        /// <summary>Gets the minute component.</summary>
        public byte Minute { get; }

        /// <summary>Gets the second component.</summary>
        public byte Second { get; }

        /// <inheritdoc />
        public override string ToString() {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0:D4}-{1:D2}-{2:D2} {3:D2}:{4:D2}:{5:D2}",
                Year,
                Month,
                Day,
                Hour,
                Minute,
                Second);
        }
    }
}
