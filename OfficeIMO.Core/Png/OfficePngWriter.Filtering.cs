using System;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private static void FilterFirstRowSub(
        byte[] rgba,
        int rowOffset,
        int stride,
        byte[] destination,
        int destinationOffset,
        System.Threading.CancellationToken cancellationToken = default,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
#if NET8_0_OR_GREATER
        if (Avx2.IsSupported && stride >= 64) {
            FilterFirstRowSubVector(rgba, rowOffset, stride, destination, destinationOffset, cancellationToken, checkpointObserver);
            return;
        }
#endif
        for (int index = 0; index < 4 && index < stride; index++) {
            destination[destinationOffset + index] = rgba[rowOffset + index];
        }
        for (int index = 4; index < stride; index++) {
            if (((index - 4) & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            destination[destinationOffset + index] = unchecked((byte)(rgba[rowOffset + index] - rgba[rowOffset + index - 4]));
        }
    }

    private static long FilterUp(
        byte[] rgba,
        int rowOffset,
        int previousRowOffset,
        int stride,
        byte[] destination,
        int destinationOffset,
        System.Threading.CancellationToken cancellationToken = default,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
#if NET8_0_OR_GREATER
        if (Avx2.IsSupported && stride >= 64) {
            return FilterUpVector(rgba, rowOffset, previousRowOffset, stride, destination, destinationOffset, cancellationToken, checkpointObserver);
        }
#endif
        long score = 0L;
        for (int index = 0; index < stride; index++) {
            if ((index & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            byte filtered = unchecked((byte)(rgba[rowOffset + index] - rgba[previousRowOffset + index]));
            destination[destinationOffset + index] = filtered;
            score += Math.Abs((int)(sbyte)filtered);
        }
        return score;
    }

    private static long FilterPaeth(
        byte[] rgba,
        int rowOffset,
        int previousRowOffset,
        int stride,
        byte[] destination,
        System.Threading.CancellationToken cancellationToken = default,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
#if NET8_0_OR_GREATER
        if (Avx2.IsSupported && Ssse3.IsSupported && stride >= 64) {
            return FilterPaethVector(rgba, rowOffset, previousRowOffset, stride, destination, cancellationToken, checkpointObserver);
        }
#endif
        long score = 0L;
        for (int index = 0; index < stride; index++) {
            if ((index & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            int left = index >= 4 ? rgba[rowOffset + index - 4] : 0;
            int above = rgba[previousRowOffset + index];
            int upperLeft = index >= 4 ? rgba[previousRowOffset + index - 4] : 0;
            byte filtered = unchecked((byte)(rgba[rowOffset + index] - PaethPredictor(left, above, upperLeft)));
            destination[index] = filtered;
            score += Math.Abs((int)(sbyte)filtered);
        }
        return score;
    }

    private static int PaethPredictor(int left, int above, int upperLeft) {
        int prediction = left + above - upperLeft;
        int distanceLeft = Math.Abs(prediction - left);
        int distanceAbove = Math.Abs(prediction - above);
        int distanceUpperLeft = Math.Abs(prediction - upperLeft);
        return distanceLeft <= distanceAbove && distanceLeft <= distanceUpperLeft
            ? left
            : distanceAbove <= distanceUpperLeft ? above : upperLeft;
    }

#if NET8_0_OR_GREATER
    // Keep the scalar checkpoint boundaries even though each arithmetic block
    // now handles several pixels. Tails use the same byte wrapping and scores.
    private static void FilterFirstRowSubVector(byte[] rgba, int rowOffset, int stride, byte[] destination, int destinationOffset,
        CancellationToken cancellationToken, Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        for (int index = 0; index < 4; index++) destination[destinationOffset + index] = rgba[rowOffset + index];
        for (int blockStart = 4; blockStart < stride;) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, stride - blockStart);
            int index = blockStart;
            for (; index <= blockEnd - 32; index += 32) {
                var current = Vector256.LoadUnsafe(ref rgba[rowOffset + index]);
                var left = Vector256.LoadUnsafe(ref rgba[rowOffset + index - 4]);
                Avx2.Subtract(current, left).StoreUnsafe(ref destination[destinationOffset + index]);
            }
            for (; index < blockEnd; index++)
                destination[destinationOffset + index] = unchecked((byte)(rgba[rowOffset + index] - rgba[rowOffset + index - 4]));
            blockStart = blockEnd;
        }
    }

    private static long FilterUpVector(byte[] rgba, int rowOffset, int previousRowOffset, int stride, byte[] destination, int destinationOffset,
        CancellationToken cancellationToken, Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        long score = 0;
        for (int blockStart = 0; blockStart < stride;) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, stride - blockStart);
            int index = blockStart;
            for (; index <= blockEnd - 32; index += 32) {
                var current = Vector256.LoadUnsafe(ref rgba[rowOffset + index]);
                var above = Vector256.LoadUnsafe(ref rgba[previousRowOffset + index]);
                var filtered = Avx2.Subtract(current, above);
                filtered.StoreUnsafe(ref destination[destinationOffset + index]);
                // Abs(-128) has the byte representation 128; unsigned SAD
                // therefore matches Math.Abs((int)(sbyte)filtered) exactly.
                var sums = Avx2.SumAbsoluteDifferences(Avx2.Abs(filtered.AsSByte()), Vector256<byte>.Zero).AsUInt64();
                score += (long)(sums.GetElement(0) + sums.GetElement(1) + sums.GetElement(2) + sums.GetElement(3));
            }
            for (; index < blockEnd; index++) {
                byte filtered = unchecked((byte)(rgba[rowOffset + index] - rgba[previousRowOffset + index]));
                destination[destinationOffset + index] = filtered;
                score += Math.Abs((int)(sbyte)filtered);
            }
            blockStart = blockEnd;
        }
        return score;
    }

    private static long FilterPaethVector(byte[] rgba, int rowOffset, int previousRowOffset, int stride, byte[] destination,
        CancellationToken cancellationToken, Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        long score = 0;
        for (int blockStart = 0; blockStart < stride;) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, stride - blockStart);
            int index = blockStart;
            for (; index < 4; index++) {
                byte filtered = unchecked((byte)(rgba[rowOffset + index] - rgba[previousRowOffset + index]));
                destination[index] = filtered;
                score += Math.Abs((int)(sbyte)filtered);
            }
            for (; index <= blockEnd - 16; index += 16) {
                var current = Vector128.LoadUnsafe(ref rgba[rowOffset + index]);
                var aboveBytes = Vector128.LoadUnsafe(ref rgba[previousRowOffset + index]);
                var left = Avx2.ConvertToVector256Int16(Vector128.LoadUnsafe(ref rgba[rowOffset + index - 4]));
                var above = Avx2.ConvertToVector256Int16(aboveBytes);
                var upperLeft = Avx2.ConvertToVector256Int16(Vector128.LoadUnsafe(ref rgba[previousRowOffset + index - 4]));
                // Distances are at most 510, so signed 16-bit arithmetic is exact.
                var distanceLeft = Avx2.Abs(Avx2.Subtract(above, upperLeft)).AsInt16();
                var distanceAbove = Avx2.Abs(Avx2.Subtract(left, upperLeft)).AsInt16();
                var distanceUpperLeft = Avx2.Abs(Avx2.Subtract(Avx2.Add(left, above), Avx2.Add(upperLeft, upperLeft))).AsInt16();
                var minimum = Avx2.Min(distanceLeft, Avx2.Min(distanceAbove, distanceUpperLeft));
                var chooseLeft = Avx2.CompareEqual(distanceLeft, minimum);
                var chooseAbove = Avx2.CompareEqual(distanceAbove, minimum);
                var prediction = Avx2.Or(Avx2.And(chooseLeft, left), Avx2.AndNot(chooseLeft,
                    Avx2.Or(Avx2.And(chooseAbove, above), Avx2.AndNot(chooseAbove, upperLeft))));
                var predictionBytes = Sse2.PackUnsignedSaturate(prediction.GetLower(), prediction.GetUpper());
                var filtered = Sse2.Subtract(current, predictionBytes);
                filtered.StoreUnsafe(ref destination[index]);
                var sums = Sse2.SumAbsoluteDifferences(Ssse3.Abs(filtered.AsSByte()), Vector128<byte>.Zero).AsUInt64();
                score += (long)(sums.GetElement(0) + sums.GetElement(1));
            }
            for (; index < blockEnd; index++) {
                int left = rgba[rowOffset + index - 4];
                int above = rgba[previousRowOffset + index];
                int upperLeft = rgba[previousRowOffset + index - 4];
                byte filtered = unchecked((byte)(rgba[rowOffset + index] - PaethPredictor(left, above, upperLeft)));
                destination[index] = filtered;
                score += Math.Abs((int)(sbyte)filtered);
            }
            blockStart = blockEnd;
        }
        return score;
    }
#endif

}
