using System;
using System.IO;
using System.Threading;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryNormalizeAdam7PngData(
        byte[] compressedData,
        int width,
        int height,
        int bitDepth,
        int colorType,
        CancellationToken cancellationToken,
        out byte[] normalizedCompressedData,
        out string? unsupportedReason) {
        normalizedCompressedData = Array.Empty<byte>();
        unsupportedReason = null;

        if (!TryGetPngChannelCount(colorType, out int channels)) {
            unsupportedReason = "Only grayscale, grayscale-alpha, indexed-color, RGB, and RGBA PNG images are currently supported.";
            return false;
        }

        int bitsPerPixel = channels * bitDepth;
        if (!TryGetPngRowByteCount(width, bitsPerPixel, out int fullRowBytes) ||
            !TryGetPngCheckedLength(fullRowBytes, height, 1, includeFilterByte: false, out int fullRowsLength) ||
            !TryDecodePngData(compressedData, cancellationToken, out byte[] decoded, out unsupportedReason)) {
            unsupportedReason ??= "PNG dimensions exceed supported limits.";
            return false;
        }

        var fullRows = new byte[fullRowsLength];
        int offset = 0;

        for (int pass = 0; pass < Adam7Passes.Length; pass++) {
            cancellationToken.ThrowIfCancellationRequested();
            Adam7Pass adam7Pass = Adam7Passes[pass];
            int passWidth = CountAdam7Samples(width, adam7Pass.XStart, adam7Pass.XStep);
            int passHeight = CountAdam7Samples(height, adam7Pass.YStart, adam7Pass.YStep);
            if (passWidth == 0 || passHeight == 0) {
                continue;
            }

            if (!TryGetPngRowByteCount(passWidth, bitsPerPixel, out int passRowBytes) ||
                !TryGetPngScanlineLength(passRowBytes, passHeight, out int passScanlineBytes)) {
                unsupportedReason = "PNG dimensions exceed supported limits.";
                return false;
            }
            if (offset + passScanlineBytes > decoded.Length) {
                unsupportedReason = "PNG image data ended before all interlaced scanlines were decoded.";
                return false;
            }

            var passScanlines = new byte[passScanlineBytes];
            CopyBytesWithCancellation(decoded, offset, passScanlines, 0, passScanlineBytes, cancellationToken);
            offset += passScanlineBytes;

            int filterBytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
            int unfilterWidth = bitDepth < 8 ? passRowBytes : passWidth;
            int unfilterBytesPerPixel = bitDepth < 8 ? 1 : filterBytesPerPixel;
            if (!TryUnfilterPngRows(passScanlines, unfilterWidth, passHeight, unfilterBytesPerPixel, cancellationToken, out var passPixels, out unsupportedReason)) {
                return false;
            }

            CopyAdam7PassPixels(passPixels, fullRows, width, bitDepth, bitsPerPixel, passWidth, passHeight, adam7Pass, cancellationToken);
        }

        if (!TryGetPngScanlineLength(fullRowBytes, height, out int normalizedRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        var normalizedRows = new byte[normalizedRowsLength];
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceRow = row * fullRowBytes;
            int targetRow = row * (fullRowBytes + 1);
            normalizedRows[targetRow] = 0;
            CopyBytesWithCancellation(fullRows, sourceRow, normalizedRows, targetRow + 1, fullRowBytes, cancellationToken);
        }

        normalizedCompressedData = DeflateZlib(normalizedRows, cancellationToken);
        return true;
    }

    private static void CopyAdam7PassPixels(
        byte[] passPixels,
        byte[] fullRows,
        int width,
        int bitDepth,
        int bitsPerPixel,
        int passWidth,
        int passHeight,
        Adam7Pass pass,
        CancellationToken cancellationToken) {
        if (!TryGetPngRowByteCount(passWidth, bitsPerPixel, out int passRowBytes) ||
            !TryGetPngRowByteCount(width, bitsPerPixel, out int fullRowBytes)) {
            return;
        }
        if (bitDepth < 8) {
            for (int y = 0; y < passHeight; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                int targetY = pass.YStart + y * pass.YStep;
                int passRow = y * passRowBytes;
                int fullRow = targetY * fullRowBytes;
                for (int x = 0; x < passWidth; x++) {
                    CheckPngLoopCancellation(PngRowLoopKind.Adam7PackedCopy, x, cancellationToken);
                    int targetX = pass.XStart + x * pass.XStep;
                    WritePackedPngSample(fullRows, fullRow, targetX, bitDepth, ReadPackedPngSample(passPixels, passRow, x, bitDepth));
                }
            }

            return;
        }

        int bytesPerPixel = bitsPerPixel / 8;
        for (int y = 0; y < passHeight; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int targetY = pass.YStart + y * pass.YStep;
            int passRow = y * passRowBytes;
            int fullRow = targetY * fullRowBytes;
            for (int x = 0; x < passWidth; x++) {
                CheckPngLoopCancellation(PngRowLoopKind.Adam7ByteCopy, x, cancellationToken);
                int targetX = pass.XStart + x * pass.XStep;
                Buffer.BlockCopy(passPixels, passRow + x * bytesPerPixel, fullRows, fullRow + targetX * bytesPerPixel, bytesPerPixel);
            }
        }
    }

    private static bool TryGetPngChannelCount(int colorType, out int channels) {
        switch (colorType) {
            case 0:
            case 3:
                channels = 1;
                return true;
            case 2:
                channels = 3;
                return true;
            case 4:
                channels = 2;
                return true;
            case 6:
                channels = 4;
                return true;
            default:
                channels = 0;
                return false;
        }
    }

    private static int CountAdam7Samples(int length, int start, int step) {
        if (length <= start) {
            return 0;
        }

        return ((length - start - 1) / step) + 1;
    }

    private static bool TryGetPngRowByteCount(int pixelCount, int bitsPerPixel, out int rowBytes) {
        rowBytes = 0;
        long bits = (long)pixelCount * bitsPerPixel;
        long bytes = (bits + 7L) / 8L;
        if (bytes > int.MaxValue || bytes > MaxPngExpandedBytes) {
            return false;
        }

        rowBytes = (int)bytes;
        return true;
    }

    private static void WritePackedPngSample(byte[] packedRows, int rowStart, int pixelIndex, int bitDepth, int sample) {
        int samplesPerByte = 8 / bitDepth;
        int targetOffset = rowStart + pixelIndex / samplesPerByte;
        int shift = (samplesPerByte - 1 - (pixelIndex % samplesPerByte)) * bitDepth;
        int mask = ((1 << bitDepth) - 1) << shift;
        packedRows[targetOffset] = (byte)((packedRows[targetOffset] & ~mask) | ((sample << shift) & mask));
    }

    private readonly struct Adam7Pass {
        internal Adam7Pass(int xStart, int yStart, int xStep, int yStep) {
            XStart = xStart;
            YStart = yStart;
            XStep = xStep;
            YStep = yStep;
        }

        internal int XStart { get; }
        internal int YStart { get; }
        internal int XStep { get; }
        internal int YStep { get; }
    }

    private static readonly Adam7Pass[] Adam7Passes = {
        new Adam7Pass(0, 0, 8, 8),
        new Adam7Pass(4, 0, 8, 8),
        new Adam7Pass(0, 4, 4, 8),
        new Adam7Pass(2, 0, 4, 4),
        new Adam7Pass(0, 2, 2, 4),
        new Adam7Pass(1, 0, 2, 2),
        new Adam7Pass(0, 1, 1, 2)
    };
}
