namespace OfficeIMO.SharedSource.Packaging;

/// <summary>Joins the CRC-32 values of consecutive byte ranges, so ranges hashed on separate threads give the whole input's CRC.</summary>
/// <remarks>Arithmetic on polynomials over GF(2) modulo the reflected CRC-32 polynomial: appending <c>n</c> bytes multiplies
/// the first CRC by x^(8n), and the second range's CRC is added.</remarks>
internal static class OfficeCrc32Combine {
    private const uint Polynomial = 0xEDB8_8320U;
    // x^(2^k) modulo the polynomial, for k = 0..31.
    private static readonly uint[] PowersOfTwo = CreatePowersOfTwo();

    /// <summary>The CRC-32 of A followed by B, from the CRC-32 of A, the CRC-32 of B and the length of B.</summary>
    internal static uint Combine(uint first, uint second, long secondLength) =>
        Multiply(PowerOfX(secondLength, 3), first) ^ second;

    private static uint Multiply(uint a, uint b) {
        uint mask = 1U << 31, product = 0;
        while (true) {
            if ((a & mask) != 0) {
                product ^= b;
                if ((a & (mask - 1)) == 0) return product;
            }
            mask >>= 1;
            b = (b & 1) != 0 ? (b >> 1) ^ Polynomial : b >> 1;
        }
    }

    /// <summary>x^(n * 2^k) modulo the polynomial.</summary>
    private static uint PowerOfX(long n, int k) {
        uint power = 1U << 31; // x^0
        while (n != 0) {
            if ((n & 1) != 0) power = Multiply(PowersOfTwo[k & 31], power);
            n >>= 1;
            k++;
        }
        return power;
    }

    private static uint[] CreatePowersOfTwo() {
        var table = new uint[32];
        uint power = 1U << 30; // x^1
        table[0] = power;
        for (int k = 1; k < table.Length; k++) table[k] = power = Multiply(power, power);
        return table;
    }
}
