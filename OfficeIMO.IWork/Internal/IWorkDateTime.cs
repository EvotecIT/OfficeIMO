namespace OfficeIMO.IWork.Internal;

/// <summary>Converts Apple epoch seconds without runtime-dependent DateTime.AddSeconds rounding.</summary>
internal static class IWorkDateTime {
    private static readonly long EpochTicks = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks;

    internal static bool TryFromAppleSeconds(double seconds, out DateTime value, bool requireExactTicks = false) {
        value = default;
        double deltaTicks = seconds * TimeSpan.TicksPerSecond;
        double roundedTicks = Math.Round(deltaTicks, MidpointRounding.AwayFromZero);
        if (double.IsNaN(roundedTicks) || double.IsInfinity(roundedTicks)
            || (requireExactTicks && deltaTicks != roundedTicks)
            || roundedTicks < -EpochTicks
            || roundedTicks > DateTime.MaxValue.Ticks - EpochTicks) {
            return false;
        }
        try {
            long ticks = checked(EpochTicks + checked((long)roundedTicks));
            if (ticks < DateTime.MinValue.Ticks || ticks > DateTime.MaxValue.Ticks) return false;
            value = new DateTime(ticks, DateTimeKind.Utc);
            return true;
        } catch (OverflowException) {
            return false;
        }
    }
}
