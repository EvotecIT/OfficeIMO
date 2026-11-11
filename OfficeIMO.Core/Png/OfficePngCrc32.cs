#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using ArmCrc32 = System.Runtime.Intrinsics.Arm.Crc32;
#endif

namespace OfficeIMO.Drawing;

/// <summary>The CRC-32 shared by PNG chunks and ZIP entries: reflected polynomial 0xEDB88320.</summary>
/// <remarks><see cref="Append(uint, byte[], int, int)"/> takes and returns the running register; start from
/// <see cref="Begin"/> and finish with <see cref="Complete"/>.</remarks>
internal static class OfficePngCrc32 {
    private const uint ReflectedPolynomial = 0xEDB88320U;
    private static readonly uint[] Table = CreateTable();

    internal static uint Begin() => 0xFFFFFFFFU;

    internal static uint Append(uint crc, byte[] data, int offset, int count) {
        ArgumentNullException.ThrowIfNull(data);
        if (offset < 0 || count < 0 || offset > data.Length - count) {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
#if NET8_0_OR_GREATER
        return Append(crc, new ReadOnlySpan<byte>(data, offset, count));
    }

    internal static uint Append(uint crc, ReadOnlySpan<byte> data) {
        if (Pclmulqdq.IsSupported && Sse2.IsSupported && data.Length >= 64) {
            crc = Fold(crc, ref data);
        } else if (ArmCrc32.Arm64.IsSupported) {
            return AppendArm(crc, data);
        }
        return AppendTable(crc, data);
#else
        int end = offset + count;
        for (; offset <= end - 8; offset += 8) {
            crc = Step8(crc, data[offset], data[offset + 1], data[offset + 2], data[offset + 3],
                data[offset + 4], data[offset + 5], data[offset + 6], data[offset + 7]);
        }
        for (; offset < end; offset++) crc = Table[(crc ^ data[offset]) & 0xFF] ^ (crc >> 8);
        return crc;
#endif
    }

    // Eight successive reflected CRC steps can be evaluated together.
    // Explicit byte assembly retains the same ordering on every target.
    private static uint Step8(uint crc, byte b0, byte b1, byte b2, byte b3, byte b4, byte b5, byte b6, byte b7) {
        uint first = crc ^ ((uint)b0 | ((uint)b1 << 8) | ((uint)b2 << 16) | ((uint)b3 << 24));
        uint second = (uint)b4 | ((uint)b5 << 8) | ((uint)b6 << 16) | ((uint)b7 << 24);
        return Table[1792 + (first & 0xFF)] ^ Table[1536 + ((first >> 8) & 0xFF)] ^
            Table[1280 + ((first >> 16) & 0xFF)] ^ Table[1024 + (first >> 24)] ^
            Table[768 + (second & 0xFF)] ^ Table[512 + ((second >> 8) & 0xFF)] ^
            Table[256 + ((second >> 16) & 0xFF)] ^ Table[second >> 24];
    }

    internal static uint Complete(uint crc) => crc ^ 0xFFFFFFFFU;

    internal static uint Compute(byte[] data, int offset, int count) =>
        Complete(Append(Begin(), data, offset, count));

#if NET8_0_OR_GREATER
    private static uint AppendTable(uint crc, ReadOnlySpan<byte> data) {
        int index = 0;
        for (; index <= data.Length - 8; index += 8) {
            crc = Step8(crc, data[index], data[index + 1], data[index + 2], data[index + 3],
                data[index + 4], data[index + 5], data[index + 6], data[index + 7]);
        }
        for (; index < data.Length; index++) {
            crc = Table[(crc ^ data[index]) & 0xFF] ^ (crc >> 8);
        }
        return crc;
    }

    // Folding works on the message as a polynomial in which the first bit read (bit 0 of byte 0) has the
    // highest degree, so bit j of a 16-byte little-endian block is the coefficient of x^(127 - j). A block A
    // that sits D bits before block C contributes A·x^D to the message, and A·x^D is congruent to
    // A_first·(x^(D+64) mod P) + A_last·(x^D mod P), where A_first and A_last are its two 64-bit halves.
    // Both products have degree below 96, so adding them to C keeps the message value modulo P.
    // A carry-less multiply of two bit-reversed 64-bit operands returns the product one position low, so each
    // constant is taken one power of x lower: x^(D+63) for the first half and x^(D-1) for the second, held in
    // the upper 32 bits of a lane in the same reflected form as the CRC register.
    private static readonly Vector128<ulong> FoldBy512 = FoldConstants(512);
    private static readonly Vector128<ulong> FoldBy128 = FoldConstants(128);

    private static Vector128<ulong> FoldConstants(int distance) =>
        Vector128.Create((ulong)PowerOfX(distance + 63) << 32, (ulong)PowerOfX(distance - 1) << 32);

    /// <summary>x^n mod P with bit 31 holding x^0, the reflected register layout.</summary>
    private static uint PowerOfX(int exponent) {
        uint value = 0x8000_0000U;
        for (int step = 0; step < exponent; step++) {
            value = (value & 1) != 0 ? (value >> 1) ^ ReflectedPolynomial : value >> 1;
        }
        return value;
    }

    /// <summary>Folds whole 16-byte blocks, leaving the unfolded tail in <paramref name="data"/>.</summary>
    /// <remarks>The register is added to the first four bytes, then the folded remainder is run through the
    /// table from a zero register, which gives the register the processed bytes would have produced.</remarks>
    private static uint Fold(uint crc, ref ReadOnlySpan<byte> data) {
        ref byte source = ref MemoryMarshal.GetReference(data);
        Vector128<ulong> a = Load(ref source, 0) ^ Vector128.CreateScalar((ulong)crc);
        Vector128<ulong> b = Load(ref source, 16);
        Vector128<ulong> c = Load(ref source, 32);
        Vector128<ulong> d = Load(ref source, 48);
        int position = 64;
        Vector128<ulong> wide = FoldBy512;
        for (; position + 64 <= data.Length; position += 64) {
            a = Multiply(a, wide) ^ Load(ref source, position);
            b = Multiply(b, wide) ^ Load(ref source, position + 16);
            c = Multiply(c, wide) ^ Load(ref source, position + 32);
            d = Multiply(d, wide) ^ Load(ref source, position + 48);
        }
        Vector128<ulong> narrow = FoldBy128;
        Vector128<ulong> folded = Multiply(Multiply(Multiply(a, narrow) ^ b, narrow) ^ c, narrow) ^ d;
        for (; position + 16 <= data.Length; position += 16) {
            folded = Multiply(folded, narrow) ^ Load(ref source, position);
        }
        data = data.Slice(position);
        Span<byte> remainder = stackalloc byte[16];
        Unsafe.WriteUnaligned(ref MemoryMarshal.GetReference(remainder), folded);
        return AppendTable(0, remainder);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<ulong> Multiply(Vector128<ulong> block, Vector128<ulong> constants) =>
        Pclmulqdq.CarrylessMultiply(block, constants, 0x00) ^ Pclmulqdq.CarrylessMultiply(block, constants, 0x11);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<ulong> Load(ref byte source, int offset) =>
        Unsafe.ReadUnaligned<Vector128<ulong>>(ref Unsafe.Add(ref source, offset));

    private static uint AppendArm(uint crc, ReadOnlySpan<byte> data) {
        int index = 0;
        for (; index + 8 <= data.Length; index += 8) {
            crc = ArmCrc32.Arm64.ComputeCrc32(crc, MemoryMarshal.Read<ulong>(data.Slice(index)));
        }
        for (; index < data.Length; index++) crc = ArmCrc32.ComputeCrc32(crc, data[index]);
        return crc;
    }
#endif

    private static uint[] CreateTable() {
        var table = new uint[8 * 256];
        for (uint index = 0; index < 256; index++) {
            uint value = index;
            for (int bit = 0; bit < 8; bit++) {
                value = (value & 1) != 0 ? ReflectedPolynomial ^ (value >> 1) : value >> 1;
            }
            table[index] = value;
        }
        for (int generation = 1; generation < 8; generation++) {
            for (int index = 0; index < 256; index++) {
                uint value = table[(generation - 1) * 256 + index];
                table[generation * 256 + index] = table[value & 0xFF] ^ (value >> 8);
            }
        }
        return table;
    }
}
