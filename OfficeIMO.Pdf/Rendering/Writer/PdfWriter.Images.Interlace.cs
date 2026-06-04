using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryNormalizeAdam7PngData(
        byte[] compressedData,
        int width,
        int height,
        int bitDepth,
        int colorType,
        out byte[] normalizedCompressedData,
        out string? unsupportedReason) {
        normalizedCompressedData = Array.Empty<byte>();
        unsupportedReason = null;

        if (!TryGetPngChannelCount(colorType, out int channels)) {
            unsupportedReason = "Only grayscale, grayscale-alpha, indexed-color, RGB, and RGBA PNG images are currently supported.";
            return false;
        }

        int bitsPerPixel = channels * bitDepth;
        int fullRowBytes = GetPngRowByteCount(width, bitsPerPixel);
        var fullRows = new byte[fullRowBytes * height];
        byte[] decoded = FlateDecoder.Decode(compressedData);
        int offset = 0;

        for (int pass = 0; pass < Adam7Passes.Length; pass++) {
            Adam7Pass adam7Pass = Adam7Passes[pass];
            int passWidth = CountAdam7Samples(width, adam7Pass.XStart, adam7Pass.XStep);
            int passHeight = CountAdam7Samples(height, adam7Pass.YStart, adam7Pass.YStep);
            if (passWidth == 0 || passHeight == 0) {
                continue;
            }

            int passRowBytes = GetPngRowByteCount(passWidth, bitsPerPixel);
            int passScanlineBytes = (passRowBytes + 1) * passHeight;
            if (offset + passScanlineBytes > decoded.Length) {
                unsupportedReason = "PNG image data ended before all interlaced scanlines were decoded.";
                return false;
            }

            var passScanlines = new byte[passScanlineBytes];
            Buffer.BlockCopy(decoded, offset, passScanlines, 0, passScanlineBytes);
            offset += passScanlineBytes;

            int filterBytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
            int unfilterWidth = bitDepth < 8 ? passRowBytes : passWidth;
            int unfilterBytesPerPixel = bitDepth < 8 ? 1 : filterBytesPerPixel;
            if (!TryUnfilterPngRows(passScanlines, unfilterWidth, passHeight, unfilterBytesPerPixel, out var passPixels, out unsupportedReason)) {
                return false;
            }

            CopyAdam7PassPixels(passPixels, fullRows, width, bitDepth, bitsPerPixel, passWidth, passHeight, adam7Pass);
        }

        var normalizedRows = new byte[(fullRowBytes + 1) * height];
        for (int row = 0; row < height; row++) {
            int sourceRow = row * fullRowBytes;
            int targetRow = row * (fullRowBytes + 1);
            normalizedRows[targetRow] = 0;
            Buffer.BlockCopy(fullRows, sourceRow, normalizedRows, targetRow + 1, fullRowBytes);
        }

        normalizedCompressedData = DeflateZlib(normalizedRows);
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
        Adam7Pass pass) {
        int passRowBytes = GetPngRowByteCount(passWidth, bitsPerPixel);
        int fullRowBytes = GetPngRowByteCount(width, bitsPerPixel);
        if (bitDepth < 8) {
            for (int y = 0; y < passHeight; y++) {
                int targetY = pass.YStart + y * pass.YStep;
                int passRow = y * passRowBytes;
                int fullRow = targetY * fullRowBytes;
                for (int x = 0; x < passWidth; x++) {
                    int targetX = pass.XStart + x * pass.XStep;
                    WritePackedPngSample(fullRows, fullRow, targetX, bitDepth, ReadPackedPngSample(passPixels, passRow, x, bitDepth));
                }
            }

            return;
        }

        int bytesPerPixel = bitsPerPixel / 8;
        for (int y = 0; y < passHeight; y++) {
            int targetY = pass.YStart + y * pass.YStep;
            int passRow = y * passRowBytes;
            int fullRow = targetY * fullRowBytes;
            for (int x = 0; x < passWidth; x++) {
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

    private static int GetPngRowByteCount(int pixelCount, int bitsPerPixel) =>
        (pixelCount * bitsPerPixel + 7) / 8;

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
