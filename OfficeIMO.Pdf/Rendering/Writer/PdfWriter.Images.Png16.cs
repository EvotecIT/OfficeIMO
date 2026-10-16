using System;
using System.IO;
using System.Threading;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryExpand16BitPng(byte[] compressedData, int width, int height, int colorType, byte[]? transparency, CancellationToken cancellationToken, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        int sourceChannels;
        int baseChannels;
        if (colorType == 0) {
            sourceChannels = 1;
            baseChannels = 1;
        } else if (colorType == 2) {
            sourceChannels = 3;
            baseChannels = 3;
        } else if (colorType == 4) {
            sourceChannels = 2;
            baseChannels = 1;
        } else if (colorType == 6) {
            sourceChannels = 4;
            baseChannels = 3;
        } else {
            unsupportedReason = "Only grayscale, grayscale-alpha, RGB, and RGBA 16-bit PNG images are currently supported.";
            return false;
        }

        bool hasIntrinsicAlpha = colorType == 4 || colorType == 6;
        int transparentGray = -1;
        int transparentRed = -1;
        int transparentGreen = -1;
        int transparentBlue = -1;
        if (transparency != null) {
            if (hasIntrinsicAlpha) {
                unsupportedReason = "PNG transparency chunks are not valid for grayscale-alpha or RGBA PNG images.";
                return false;
            }

            int requiredTransparencyLength = colorType == 0 ? 2 : 6;
            if (transparency.Length < requiredTransparencyLength) {
                unsupportedReason = colorType == 0
                    ? "Grayscale PNG transparency chunk is invalid."
                    : "RGB PNG transparency chunk is invalid.";
                return false;
            }

            transparentGray = colorType == 0 ? ReadUInt16BigEndian(transparency, 0) : -1;
            transparentRed = colorType == 2 ? ReadUInt16BigEndian(transparency, 0) : -1;
            transparentGreen = colorType == 2 ? ReadUInt16BigEndian(transparency, 2) : -1;
            transparentBlue = colorType == 2 ? ReadUInt16BigEndian(transparency, 4) : -1;
        }

        if (!TryDecodePngData(compressedData, cancellationToken, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        int sourceBytesPerPixel = sourceChannels * 2;
        if (!TryGetPngCheckedLength(width, height, sourceBytesPerPixel, includeFilterByte: true, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length != expectedLength) {
            unsupportedReason = "PNG image data ended before all 16-bit scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, width, height, sourceBytesPerPixel, cancellationToken, out var rawPixels, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, baseChannels * 2, includeFilterByte: true, out int baseRowsLength) ||
            !TryGetPngCheckedLength(width, height, 2, includeFilterByte: true, out int alphaRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[baseRowsLength];
        byte[]? alphaRows = hasIntrinsicAlpha || transparency != null ? new byte[alphaRowsLength] : null;

        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int baseRowStart = row * (1 + width * baseChannels * 2);
            int alphaRowStart = row * (1 + width * 2);
            baseRows[baseRowStart] = 0;
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * width * sourceBytesPerPixel;
            for (int pixel = 0; pixel < width; pixel++) {
                CheckPngLoopCancellation(PngRowLoopKind.SixteenBitExpansion, pixel, cancellationToken);
                int sourcePixel = sourceRowStart + pixel * sourceBytesPerPixel;
                int basePixel = baseRowStart + 1 + pixel * baseChannels * 2;

                if (colorType == 0 || colorType == 4) {
                    int gray = ReadUInt16BigEndian(rawPixels, sourcePixel);
                    WritePng16Sample(baseRows, basePixel, gray);
                    if (alphaRows != null) {
                        int alpha = hasIntrinsicAlpha
                            ? ReadUInt16BigEndian(rawPixels, sourcePixel + 2)
                            : gray == transparentGray ? 0 : 65535;
                        WritePng16Sample(alphaRows, alphaRowStart + 1 + pixel * 2, alpha);
                    }
                } else {
                    int red = ReadUInt16BigEndian(rawPixels, sourcePixel);
                    int green = ReadUInt16BigEndian(rawPixels, sourcePixel + 2);
                    int blue = ReadUInt16BigEndian(rawPixels, sourcePixel + 4);
                    WritePng16Sample(baseRows, basePixel, red);
                    WritePng16Sample(baseRows, basePixel + 2, green);
                    WritePng16Sample(baseRows, basePixel + 4, blue);
                    if (alphaRows != null) {
                        int alpha = hasIntrinsicAlpha
                            ? ReadUInt16BigEndian(rawPixels, sourcePixel + 6)
                            : red == transparentRed && green == transparentGreen && blue == transparentBlue ? 0 : 65535;
                        WritePng16Sample(alphaRows, alphaRowStart + 1 + pixel * 2, alpha);
                    }
                }
            }
        }

        string colorSpace = baseChannels == 1 ? "/DeviceGray" : "/DeviceRGB";
        image = new PdfImageStream {
            Data = DeflateZlib(baseRows, cancellationToken),
            PixelWidth = width,
            PixelHeight = height,
            BitsPerComponent = 16,
            DictionarySuffix = BuildPngPredictorDictionarySuffix(colorSpace, baseChannels, width, 16)
        };

        if (alphaRows != null) {
            image.SoftMask = new PdfImageStream {
                Data = DeflateZlib(alphaRows, cancellationToken),
                PixelWidth = width,
                PixelHeight = height,
                BitsPerComponent = 16,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width, 16)
            };
        }

        return true;
    }

    private static void WritePng16Sample(byte[] bytes, int offset, int sample) {
        bytes[offset] = (byte)(sample >> 8);
        bytes[offset + 1] = (byte)sample;
    }

    /// <summary>
    /// PDF 1.5 introduced 16-bit image samples. Older targets retain the existing nearest-value
    /// reduction, round(sample / 257), for both color and alpha without mutating cached preparation.
    /// </summary>
    internal static PdfImageStream PrepareImageStreamForVersion(PdfImageStream source, PdfFileVersion version,
        CancellationToken cancellationToken = default) {
        if (version >= PdfFileVersion.Pdf15 || source.BitsPerComponent != 16) return source;
        PdfImageStream image = source.CloneForWrite();
        int channels = source.DictionarySuffix.Contains("/DeviceGray") ? 1 : 3;
        if (!TryDecodePngData(source.Data, cancellationToken, out byte[] decoded, out string? reason) ||
            !TryUnfilterPngRows(decoded, source.PixelWidth, source.PixelHeight, channels * 2,
                cancellationToken, out byte[] samples, out reason) ||
            !TryGetPngCheckedLength(source.PixelWidth, source.PixelHeight, channels, true, out int length)) {
            throw new NotSupportedException(reason ?? "16-bit image reduction exceeds supported limits.");
        }
        byte[] rows = new byte[length];
        int stride = source.PixelWidth * channels;
        for (int row = 0; row < source.PixelHeight; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int sample = 0; sample < stride; sample++) {
                CheckPngLoopCancellation(PngRowLoopKind.SixteenBitExpansion, sample, cancellationToken);
                rows[row * (stride + 1) + 1 + sample] =
                    Scale16BitSampleToByte(ReadUInt16BigEndian(samples, (row * stride + sample) * 2));
            }
        }
        image.Data = DeflateZlib(rows, cancellationToken);
        image.BitsPerComponent = 8;
        image.DictionarySuffix = image.DictionarySuffix.Replace("/BitsPerComponent 16", "/BitsPerComponent 8");
        if (image.SoftMask != null) image.SoftMask = PrepareImageStreamForVersion(image.SoftMask, version, cancellationToken);
        return image;
    }

    private static byte Scale16BitSampleToByte(int sample) =>
        (byte)((sample + 128) / 257);
}
