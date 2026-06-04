using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryExpand16BitPng(byte[] compressedData, int width, int height, int colorType, byte[]? transparency, out PdfImageStream image, out string? unsupportedReason) {
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

        byte[] decoded = FlateDecoder.Decode(compressedData);
        int sourceBytesPerPixel = sourceChannels * 2;
        int expectedRowLength = 1 + width * sourceBytesPerPixel;
        if (decoded.Length < expectedRowLength * height) {
            unsupportedReason = "PNG image data ended before all 16-bit scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, width, height, sourceBytesPerPixel, out var rawPixels, out unsupportedReason)) {
            return false;
        }

        byte[] baseRows = new byte[(1 + width * baseChannels) * height];
        byte[]? alphaRows = hasIntrinsicAlpha || transparency != null ? new byte[(1 + width) * height] : null;

        for (int row = 0; row < height; row++) {
            int baseRowStart = row * (1 + width * baseChannels);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * width * sourceBytesPerPixel;
            for (int pixel = 0; pixel < width; pixel++) {
                int sourcePixel = sourceRowStart + pixel * sourceBytesPerPixel;
                int basePixel = baseRowStart + 1 + pixel * baseChannels;

                if (colorType == 0 || colorType == 4) {
                    int gray = ReadUInt16BigEndian(rawPixels, sourcePixel);
                    baseRows[basePixel] = Scale16BitSampleToByte(gray);
                    if (alphaRows != null) {
                        int alpha = hasIntrinsicAlpha
                            ? ReadUInt16BigEndian(rawPixels, sourcePixel + 2)
                            : gray == transparentGray ? 0 : 65535;
                        alphaRows[alphaRowStart + 1 + pixel] = Scale16BitSampleToByte(alpha);
                    }
                } else {
                    int red = ReadUInt16BigEndian(rawPixels, sourcePixel);
                    int green = ReadUInt16BigEndian(rawPixels, sourcePixel + 2);
                    int blue = ReadUInt16BigEndian(rawPixels, sourcePixel + 4);
                    baseRows[basePixel] = Scale16BitSampleToByte(red);
                    baseRows[basePixel + 1] = Scale16BitSampleToByte(green);
                    baseRows[basePixel + 2] = Scale16BitSampleToByte(blue);
                    if (alphaRows != null) {
                        int alpha = hasIntrinsicAlpha
                            ? ReadUInt16BigEndian(rawPixels, sourcePixel + 6)
                            : red == transparentRed && green == transparentGreen && blue == transparentBlue ? 0 : 65535;
                        alphaRows[alphaRowStart + 1 + pixel] = Scale16BitSampleToByte(alpha);
                    }
                }
            }
        }

        string colorSpace = baseChannels == 1 ? "/DeviceGray" : "/DeviceRGB";
        image = new PdfImageStream {
            Data = DeflateZlib(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix(colorSpace, baseChannels, width)
        };

        if (alphaRows != null) {
            image.SoftMask = new PdfImageStream {
                Data = DeflateZlib(alphaRows),
                PixelWidth = width,
                PixelHeight = height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
            };
        }

        return true;
    }

    private static byte Scale16BitSampleToByte(int sample) =>
        (byte)((sample + 128) / 257);
}
