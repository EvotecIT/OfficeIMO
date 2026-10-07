#if NET8_0_OR_GREATER
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngReader {
    // The three channels depend on the previous pixel independently. Reads and
    // writes touch exactly three bytes, including the last pixel of a row.
    private static void UnfilterRgbAverage(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        ref byte upper = ref MemoryMarshal.GetReference(previous);
        var left = Vector128<byte>.Zero;
        var one = Vector128.Create((byte)1);
        for (int blockStart = 0; blockStart < current.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            // Whole pixels keep each cancellation interval within 3072 bytes.
            int blockEnd = blockStart + Math.Min(3072, current.Length - blockStart);
            for (int index = blockStart; index < blockEnd; index += 3) {
                var above = ReadRgbPixel(ref Unsafe.Add(ref upper, index));
                // PAVG rounds up; removing the odd-sum bit retains PNG's floor.
                var prediction = Sse2.Subtract(Sse2.Average(left, above),
                    Sse2.And(Sse2.Xor(left, above), one));
                left = Sse2.Add(ReadRgbPixel(ref Unsafe.Add(ref data, index)), prediction);
                WriteRgbPixel(ref Unsafe.Add(ref data, index), left);
            }
            blockStart = blockEnd;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<byte> ReadRgbPixel(ref byte pixel) =>
        Vector128.CreateScalar((uint)Unsafe.ReadUnaligned<ushort>(ref pixel)
            | ((uint)Unsafe.Add(ref pixel, 2) << 16)).AsByte();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void WriteRgbPixel(ref byte pixel, Vector128<byte> value) {
        uint channels = value.AsUInt32().ToScalar();
        Unsafe.WriteUnaligned(ref pixel, (ushort)channels);
        Unsafe.Add(ref pixel, 2) = (byte)(channels >> 16);
    }
}
#endif
