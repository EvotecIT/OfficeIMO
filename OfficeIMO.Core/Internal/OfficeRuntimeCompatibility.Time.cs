namespace OfficeIMO.Internal {
    internal static partial class OfficeRuntimeCompatibility {
        // .NET Framework rounds fractional DateTime.Add* and TimeSpan.From* arguments to whole milliseconds;
        // .NET 8 and later keep the sub-millisecond ticks. These follow the .NET 8 arithmetic on every framework,
        // so values read from documents get the same ticks wherever they are read.

        /// <summary>Adds fractional seconds to a date as <c>DateTime.AddSeconds</c> does on .NET 8 and later.</summary>
        internal static DateTime AddSeconds(DateTime value, double seconds) =>
            value.AddTicks(UnitsToTicks(seconds, TimeSpan.MaxValue.Ticks / TimeSpan.TicksPerSecond, TimeSpan.TicksPerSecond));

        /// <summary>Adds fractional seconds to a date and offset as <c>DateTimeOffset.AddSeconds</c> does on .NET 8 and later.</summary>
        internal static DateTimeOffset AddSeconds(DateTimeOffset value, double seconds) =>
            value.AddTicks(UnitsToTicks(seconds, TimeSpan.MaxValue.Ticks / TimeSpan.TicksPerSecond, TimeSpan.TicksPerSecond));

        /// <summary>Adds fractional days to a date as <c>DateTime.AddDays</c> does on .NET 8 and later.</summary>
        internal static DateTime AddDays(DateTime value, double days) =>
            value.AddTicks(UnitsToTicks(days, TimeSpan.MaxValue.Ticks / TimeSpan.TicksPerDay, TimeSpan.TicksPerDay));

        /// <summary>Converts fractional seconds to a duration as <c>TimeSpan.FromSeconds</c> does on .NET 8 and later.</summary>
        internal static TimeSpan FromSeconds(double seconds) {
            if (double.IsNaN(seconds)) throw new ArgumentException("The duration cannot be NaN.", nameof(seconds));
            double ticks = seconds * TimeSpan.TicksPerSecond;
            if (ticks > long.MaxValue || ticks < long.MinValue) throw new OverflowException("The duration is too long for a TimeSpan.");
            return ticks == long.MaxValue ? TimeSpan.MaxValue : new TimeSpan((long)ticks);
        }

        // The whole units and the fraction are converted separately and each truncated toward zero.
        private static long UnitsToTicks(double value, long maximumUnits, long ticksPerUnit) {
            if (double.IsNaN(value) || Math.Abs(value) > maximumUnits) {
                throw new ArgumentOutOfRangeException(nameof(value), "The added value is outside the range of DateTime.");
            }
            double whole = Math.Truncate(value);
            return (long)whole * ticksPerUnit + (long)((value - whole) * ticksPerUnit);
        }
    }
}
