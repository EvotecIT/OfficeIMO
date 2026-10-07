using System.Globalization;

namespace OfficeIMO.Epub;

/// <summary>Exact, bounded SMIL clock handling for media-overlay duration metadata.</summary>
internal static class EpubSmilClock {
    internal static string Format(TimeSpan value) => ((decimal)value.Ticks / TimeSpan.TicksPerSecond).ToString("0.#######", CultureInfo.InvariantCulture) + "s";

    internal static TimeSpan Parse(string value) {
        // SMIL full/partial clocks and timecount units. No signs, exponents or culture-dependent syntax.
        string text = value.Trim();
        if (text.Length == 0 || text.Length > 64) throw Invalid();
        decimal seconds;
        string[] parts = text.Split(':');
        if (parts.Length == 2 || parts.Length == 3) {
            string secondPart = parts[parts.Length - 1];
            int secondDigits = secondPart.IndexOf('.') < 0 ? secondPart.Length : secondPart.IndexOf('.');
            if (secondDigits != 2 || parts[parts.Length - 2].Length != 2) throw Invalid();
            decimal tail = Number(parts[parts.Length - 1], fraction: true);
            decimal minutes = Number(parts[parts.Length - 2], fraction: false);
            if (tail >= 60 || minutes >= 60) throw Invalid();
            decimal hours = parts.Length == 3 ? Number(parts[0], fraction: false) : 0;
            try { seconds = checked(hours * 3600 + minutes * 60 + tail); }
            catch (OverflowException) { throw Invalid(); }
        } else if (parts.Length == 1) {
            decimal multiplier = 1;
            if (text.EndsWith("ms", StringComparison.Ordinal)) { multiplier = 0.001m; text = text.Substring(0, text.Length - 2); }
            else if (text.EndsWith("min", StringComparison.Ordinal)) { multiplier = 60; text = text.Substring(0, text.Length - 3); }
            else if (text.EndsWith("h", StringComparison.Ordinal)) { multiplier = 3600; text = text.Substring(0, text.Length - 1); }
            else if (text.EndsWith("s", StringComparison.Ordinal)) text = text.Substring(0, text.Length - 1);
            try { seconds = checked(Number(text, fraction: true) * multiplier); }
            catch (OverflowException) { throw Invalid(); }
        } else throw Invalid();
        decimal ticks;
        try { ticks = checked(seconds * TimeSpan.TicksPerSecond); }
        catch (OverflowException) { throw Invalid(); }
        if (ticks > long.MaxValue || ticks != decimal.Truncate(ticks)) throw Invalid();
        return TimeSpan.FromTicks((long)ticks);
    }

    private static decimal Number(string text, bool fraction) {
        int dot = text.IndexOf('.');
        if (text.Length == 0 || text.Length > 28 || dot == 0 || dot == text.Length - 1 || !fraction && dot >= 0 ||
            text.Any(c => (c < '0' || c > '9') && c != '.') ||
            !decimal.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal number)) throw Invalid();
        return number;
    }
    private static InvalidDataException Invalid() => new InvalidDataException("Existing media duration requires a nonnegative SMIL clock representable exactly as TimeSpan.");
}
