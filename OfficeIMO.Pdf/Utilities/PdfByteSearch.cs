namespace OfficeIMO.Pdf;

/// <summary>
/// Byte searches over whole PDF buffers. Modern targets use the vectorized span search; older targets have no span
/// search, and <c>Array.IndexOf</c> there reads one byte per step behind a generic comparer, so they test eight bytes
/// per step instead.
/// </summary>
internal static class PdfByteSearch {
    /// <summary>Returns the index of the first <paramref name="value"/> in the range, or -1.</summary>
    internal static int IndexOf(byte[] bytes, byte value, int start, int count) {
        Guard.NotNull(bytes, nameof(bytes));
        if (start < 0 || count < 0 || start > bytes.Length - count) throw new ArgumentOutOfRangeException(nameof(start));
#if NET8_0_OR_GREATER
        int relative = bytes.AsSpan(start, count).IndexOf(value);
        return relative < 0 ? -1 : start + relative;
#else
        unsafe {
            fixed (byte* first = bytes) {
                byte* current = first + start;
                byte* end = current + count;
                const ulong Ones = 0x0101010101010101UL;
                const ulong Highs = 0x8080808080808080UL;
                ulong pattern = Ones * value;
                while (end - current >= 8) {
                    // A zero byte in x marks a match; the lowest flagged byte is always a real match.
                    ulong x = *(ulong*)current ^ pattern;
                    ulong found = (x - Ones) & ~x & Highs;
                    if (found != 0) return (int)(current - first) + LowestFlaggedByte(found);
                    current += 8;
                }
                while (current < end) {
                    if (*current == value) return (int)(current - first);
                    current++;
                }
                return -1;
            }
        }
#endif
    }

#if !NET8_0_OR_GREATER
    private static int LowestFlaggedByte(ulong found) {
        int index = 0;
        if ((found & 0xFFFFFFFFUL) == 0) { index += 4; found >>= 32; }
        if ((found & 0xFFFFUL) == 0) { index += 2; found >>= 16; }
        if ((found & 0xFFUL) == 0) index += 1;
        return index;
    }
#endif
}
