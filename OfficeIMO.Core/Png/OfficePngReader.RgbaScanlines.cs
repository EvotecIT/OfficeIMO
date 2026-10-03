#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngReader {
    // RGBA8 rows can be reconstructed in the final owned buffer. Earlier rows
    // supply the immutable above pixels, avoiding row copies and clearing.
    private static void DecodeRgbaScanlines(PngPayload payload, OfficeRasterImage result,
        CancellationToken cancellationToken) {
        var zeroRow = new byte[payload.Stride];
        byte[] pixels = result.PixelBuffer;
        int sourceOffset = 0;
        for (int y = 0; y < payload.Height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            int filter = payload.Scanlines[sourceOffset++];
            int targetOffset = y * payload.Stride;
            CopyBytes(payload.Scanlines, sourceOffset, pixels, targetOffset, payload.Stride, cancellationToken);
            sourceOffset += payload.Stride;
            ReadOnlySpan<byte> previous = y == 0
                ? zeroRow
                : pixels.AsSpan(targetOffset - payload.Stride, payload.Stride);
            UnfilterRgba8(pixels.AsSpan(targetOffset, payload.Stride), previous, filter, cancellationToken);
        }
    }

    private static void UnfilterRgba8(Span<byte> current, ReadOnlySpan<byte> previous, int filter,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (filter) {
            case 0:
                return;
            case 1:
                UnfilterRgbaSub(current, cancellationToken);
                return;
            case 2:
                UnfilterRgbaUp(current, previous, cancellationToken);
                return;
            case 3:
                UnfilterRgbaAverage(current, previous, cancellationToken);
                return;
            case 4:
                UnfilterRgbaPaeth(current, previous, cancellationToken);
                return;
            default:
                throw new InvalidDataException("Unsupported PNG filter.");
        }
    }

    private static void UnfilterRgbaSub(Span<byte> current, CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        uint carry = 0;
        for (int blockStart = 0; blockStart < current.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, current.Length - blockStart);
            int index = blockStart;
            for (; index <= blockEnd - 16; index += 16) {
                var value = Vector128.LoadUnsafe(ref Unsafe.Add(ref data, index));
                // Four independent modulo-256 prefix sums, one per channel.
                value = Sse2.Add(value, Sse2.ShiftLeftLogical128BitLane(value, 4));
                value = Sse2.Add(value, Sse2.ShiftLeftLogical128BitLane(value, 8));
                value = Sse2.Add(value, Vector128.Create(carry).AsByte());
                value.StoreUnsafe(ref Unsafe.Add(ref data, index));
                carry = value.AsUInt32().GetElement(3);
            }
            for (; index < blockEnd; index += 4) {
                var value = Sse2.Add(ReadRgbaPixel(ref Unsafe.Add(ref data, index)), Vector128.Create(carry).AsByte());
                WriteRgbaPixel(ref Unsafe.Add(ref data, index), value);
                carry = value.AsUInt32().ToScalar();
            }
            blockStart = blockEnd;
        }
    }

    private static void UnfilterRgbaUp(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        ref byte above = ref MemoryMarshal.GetReference(previous);
        for (int blockStart = 0; blockStart < current.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, current.Length - blockStart);
            int index = blockStart;
            for (; index <= blockEnd - 16; index += 16) {
                var value = Vector128.LoadUnsafe(ref Unsafe.Add(ref data, index));
                var upper = Vector128.LoadUnsafe(ref Unsafe.Add(ref above, index));
                Sse2.Add(value, upper).StoreUnsafe(ref Unsafe.Add(ref data, index));
            }
            for (; index < blockEnd; index++) {
                current[index] = unchecked((byte)(current[index] + previous[index]));
            }
            blockStart = blockEnd;
        }
    }

    private static void UnfilterRgbaAverage(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        ref byte upper = ref MemoryMarshal.GetReference(previous);
        var left = Vector128<byte>.Zero;
        var one = Vector128.Create((byte)1);
        for (int index = 0; index < current.Length; index += 4) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            var above = ReadRgbaPixel(ref Unsafe.Add(ref upper, index));
            // PAVG rounds up; subtract the odd-sum bit to retain PNG's floor.
            var prediction = Sse2.Subtract(Sse2.Average(left, above), Sse2.And(Sse2.Xor(left, above), one));
            left = Sse2.Add(ReadRgbaPixel(ref Unsafe.Add(ref data, index)), prediction);
            WriteRgbaPixel(ref Unsafe.Add(ref data, index), left);
        }
    }

    private static void UnfilterRgbaPaeth(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        ref byte upper = ref MemoryMarshal.GetReference(previous);
        var zero = Vector128<byte>.Zero;
        var left = Vector128<short>.Zero;
        var upperLeft = Vector128<short>.Zero;
        for (int index = 0; index < current.Length; index += 4) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            var above = Sse2.UnpackLow(ReadRgbaPixel(ref Unsafe.Add(ref upper, index)), zero).AsInt16();
            // The four unsigned channels widen to signed shorts. Predictor
            // distances stay within 510 and preserve left/above/diagonal ties.
            var distanceLeft = Ssse3.Abs(Sse2.Subtract(above, upperLeft)).AsInt16();
            var distanceAbove = Ssse3.Abs(Sse2.Subtract(left, upperLeft)).AsInt16();
            var distanceDiagonal = Ssse3.Abs(Sse2.Subtract(Sse2.Add(left, above), Sse2.Add(upperLeft, upperLeft))).AsInt16();
            var minimum = Sse2.Min(distanceLeft, Sse2.Min(distanceAbove, distanceDiagonal));
            var chooseLeft = Sse2.CompareEqual(distanceLeft, minimum);
            var chooseAbove = Sse2.CompareEqual(distanceAbove, minimum);
            var prediction = Sse2.Or(Sse2.And(chooseLeft, left), Sse2.AndNot(chooseLeft,
                Sse2.Or(Sse2.And(chooseAbove, above), Sse2.AndNot(chooseAbove, upperLeft))));
            var predictedBytes = Sse2.PackUnsignedSaturate(prediction, Vector128<short>.Zero);
            var value = Sse2.Add(ReadRgbaPixel(ref Unsafe.Add(ref data, index)), predictedBytes);
            WriteRgbaPixel(ref Unsafe.Add(ref data, index), value);
            left = Sse2.UnpackLow(value, zero).AsInt16();
            upperLeft = above;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<byte> ReadRgbaPixel(ref byte pixel) =>
        Vector128.CreateScalar(Unsafe.ReadUnaligned<uint>(ref pixel)).AsByte();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void WriteRgbaPixel(ref byte pixel, Vector128<byte> value) =>
        Unsafe.WriteUnaligned(ref pixel, value.AsUInt32().ToScalar());
}
#endif
