using System;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // PNG and PDF PNG predictors share byte-oriented rows, including packed gray samples.
    // Validate the complete payload before retaining it as a PDF image stream.
    private static bool TryValidatePngPassThroughData(byte[] compressedData, int width, int height, int colors, int bitDepth, CancellationToken cancellationToken, out string? unsupportedReason) {
        if (!TryDecodePngData(compressedData, cancellationToken, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngRowByteCount(width, colors * bitDepth, out int rowBytes) ||
            !TryGetPngScanlineLength(rowBytes, height, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length != expectedLength) {
            unsupportedReason = "PNG image data length does not match the expected scanline size.";
            return false;
        }

        int rowLength = rowBytes + 1;
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int filter = decoded[row * rowLength];
            if (filter > 4) {
                unsupportedReason = "Unsupported PNG scanline filter: " + filter.ToString(CultureInfo.InvariantCulture) + ".";
                return false;
            }
        }

        return true;
    }

    private static bool TryExpandPackedGrayscalePng(byte[] compressedData, int width, int height, int bitDepth, byte[]? transparency, CancellationToken cancellationToken, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        int maxSample = (1 << bitDepth) - 1;
        int transparentSample = -1;
        if (transparency != null) {
            if (transparency.Length < 2) {
                unsupportedReason = "Grayscale PNG transparency chunk is invalid.";
                return false;
            }

            transparentSample = ReadUInt16BigEndian(transparency, 0);
            if (transparentSample > maxSample) {
                unsupportedReason = "Grayscale PNG transparency value exceeds the image bit depth.";
                return false;
            }
        }

        if (!TryDecodePngData(compressedData, cancellationToken, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngRowByteCount(width, bitDepth, out int packedRowBytes) ||
            !TryGetPngScanlineLength(packedRowBytes, height, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG image data ended before all grayscale scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, packedRowBytes, height, 1, cancellationToken, out var packedRows, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int grayscaleRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[grayscaleRowsLength];
        byte[]? alphaRows = transparency != null ? new byte[grayscaleRowsLength] : null;
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int baseRowStart = row * (1 + width);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * packedRowBytes;
            for (int pixel = 0; pixel < width; pixel++) {
                CheckPngLoopCancellation(PngRowLoopKind.PackedGrayscale, pixel, cancellationToken);
                int sample = ReadPackedPngSample(packedRows, sourceRowStart, pixel, bitDepth);
                int targetOffset = baseRowStart + 1 + pixel;
                baseRows[targetOffset] = ScalePackedSampleToByte(sample, maxSample);
                if (alphaRows != null) {
                    alphaRows[alphaRowStart + 1 + pixel] = sample == transparentSample ? (byte)0 : (byte)255;
                }
            }
        }

        image = new PdfImageStream {
            Data = DeflateZlib(baseRows, cancellationToken),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
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

    private static string BuildPngPredictorDictionarySuffix(string colorSpace, int colors, int width, int bitDepth = 8) =>
        " /ColorSpace " + colorSpace +
        " /BitsPerComponent " + bitDepth.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors " +
        colors.ToString(CultureInfo.InvariantCulture) +
        " /BitsPerComponent " + bitDepth.ToString(CultureInfo.InvariantCulture) + " /Columns " +
        width.ToString(CultureInfo.InvariantCulture) +
        " >>";

}
