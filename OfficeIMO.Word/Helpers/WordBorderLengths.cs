using System;

namespace OfficeIMO.Word;

/// <summary>Converts physical border lengths at the WordprocessingML storage boundary.</summary>
internal static class WordBorderLengths {
    internal static OfficeLength? FromSize(uint? value) =>
        value.HasValue ? OfficeLength.Points(value.Value / 8D) : (OfficeLength?)null;

    internal static OfficeLength? FromSpace(uint? value) =>
        value.HasValue ? OfficeLength.Emus((long)value.Value * 12700) : (OfficeLength?)null;

    internal static uint? ToSize(OfficeLength? value) => ToNative(value, 8);

    internal static uint? ToSpace(OfficeLength? value) => ToNative(value, 1);

    // Preserve zero, unset attributes and the existing UInt32 storage range. Reject
    // negative EMUs before rounding so a tiny negative length cannot become zero.
    private static uint? ToNative(OfficeLength? value, int unitsPerPoint) {
        if (!value.HasValue) return null;
        if (value.Value < OfficeLength.Zero)
            throw new ArgumentOutOfRangeException(nameof(value), "The length cannot be negative.");
        return checked((uint)decimal.Round(value.Value.ToEmus() * (decimal)unitsPerPoint / 12700m,
            0, MidpointRounding.AwayFromZero));
    }
}
