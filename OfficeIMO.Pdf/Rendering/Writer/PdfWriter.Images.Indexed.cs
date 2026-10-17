using System;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryExpandIndexedPng(byte[] compressedData, int width, int height, int bitDepth, byte[]? palette, byte[]? paletteAlpha, CancellationToken cancellationToken, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        if (palette == null || palette.Length == 0) {
            unsupportedReason = "Indexed-color PNG images require a PLTE palette chunk.";
            return false;
        }

        int paletteEntries = palette.Length / 3;
        if (paletteAlpha != null && paletteAlpha.Length > paletteEntries) {
            unsupportedReason = "Indexed-color PNG transparency has more entries than the PLTE palette.";
            return false;
        }

        if (!TryDecodePngData(compressedData, cancellationToken, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngRowByteCount(width, bitDepth, out int packedRowBytes) ||
            !TryGetPngScanlineLength(packedRowBytes, height, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length != expectedLength) {
            unsupportedReason = "PNG image data ended before all indexed-color scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, packedRowBytes, height, 1, cancellationToken, out var packedRows, out unsupportedReason)) {
            return false;
        }

        if (paletteEntries > (1 << bitDepth) ||
            !TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int indexedAlphaRowsLength)) {
            unsupportedReason = "PNG palette or dimensions exceed supported limits.";
            return false;
        }

        byte[]? alphaRows = HasPaletteTransparency(paletteAlpha) ? new byte[indexedAlphaRowsLength] : null;
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int alphaRowStart = row * (1 + width);
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * packedRowBytes;
            for (int pixel = 0; pixel < width; pixel++) {
                CheckPngLoopCancellation(PngRowLoopKind.IndexedExpansion, pixel, cancellationToken);
                int paletteIndex = ReadPackedPngSample(packedRows, sourceRowStart, pixel, bitDepth);
                if (paletteIndex >= paletteEntries) {
                    unsupportedReason = "Indexed-color PNG pixel references a palette entry that does not exist.";
                    return false;
                }

                if (alphaRows != null) {
                    alphaRows[alphaRowStart + 1 + pixel] =
                        paletteAlpha != null && paletteIndex < paletteAlpha.Length ? paletteAlpha[paletteIndex] : (byte)255;
                }
            }
        }

        image = new PdfImageStream {
            Data = compressedData,
            PixelWidth = width,
            PixelHeight = height,
            BitsPerComponent = bitDepth,
            IndexedPalette = palette,
            DictionarySuffix = BuildPngPredictorDictionarySuffix("[/Indexed /DeviceRGB " +
                (paletteEntries - 1).ToString(CultureInfo.InvariantCulture) + " <" + ToHex(palette) + ">]", 1, width, bitDepth)
        };
        if (alphaRows != null) {
            image.SoftMask = new PdfImageStream {
                Data = DeflateZlib(alphaRows, cancellationToken),
                PixelWidth = width,
                PixelHeight = height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
            };
        }

        return true;
    }

}
