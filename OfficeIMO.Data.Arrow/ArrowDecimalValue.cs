using System.Numerics;

namespace OfficeIMO.Data.Arrow;

/// <summary>Checks the final Arrow coefficient and removes representation-only excess scale.</summary>
internal static class ArrowDecimalValue {
    private static readonly BigInteger[] PowersOfTen = CreatePowersOfTen();

    internal static decimal NormalizeExact(decimal value, int ordinal, int scale, int precision, BigInteger precisionLimit) {
        Span<int> bits = stackalloc int[4];
        decimal.GetBits(value, bits);
        int sourceScale = (bits[3] >> 16) & 0xff;
        if (sourceScale > scale) {
            decimal normalized = decimal.Round(value, scale, MidpointRounding.ToEven);
            if (normalized != value) {
                throw new InvalidDataException(
                    $"Decimal value in column {ordinal} cannot be represented exactly with Arrow scale {scale}. " +
                    "Increase ArrowReadOptions.DecimalScale to preserve the value.");
            }
            value = normalized;
            decimal.GetBits(value, bits);
            sourceScale = (bits[3] >> 16) & 0xff;
        }

        BigInteger coefficient = new BigInteger((uint)bits[0]) +
            (new BigInteger((uint)bits[1]) << 32) + (new BigInteger((uint)bits[2]) << 64);
        coefficient *= PowersOfTen[scale - sourceScale];
        if (coefficient >= precisionLimit) {
            throw new InvalidDataException(
                $"Decimal value in column {ordinal} exceeds Arrow precision {precision} at scale {scale}.");
        }
        return value;
    }

    private static BigInteger[] CreatePowersOfTen() {
        var powers = new BigInteger[39];
        powers[0] = BigInteger.One;
        for (int i = 1; i < powers.Length; i++) powers[i] = powers[i - 1] * 10;
        return powers;
    }
}
