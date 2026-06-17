using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const long MaxPngPixelCount = 100_000_000L;
    private const long MaxPngExpandedBytes = 256L * 1024L * 1024L;
    private const int MaxPngDecodedBytes = 256 * 1024 * 1024;

    internal sealed class PdfImageStream {
        public byte[] Data { get; set; } = Array.Empty<byte>();
        public string DictionarySuffix { get; set; } = string.Empty;
        public int PixelWidth { get; set; }
        public int PixelHeight { get; set; }
        public PdfImageStream? SoftMask { get; set; }
    }

    internal static bool TryGetPngImageData(byte[] data, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        if (data.Length < 33 || !IsPng(data)) {
            unsupportedReason = "Bytes are not a PNG image.";
            return false;
        }

        int offset = 8;
        int width = 0;
        int height = 0;
        int bitDepth = 0;
        int colorType = -1;
        int compression = 0;
        int filter = 0;
        int interlace = 0;
        byte[]? palette = null;
        byte[]? transparency = null;
        var idat = new MemoryStream();

        while (offset + 12 <= data.Length) {
            int length = ReadInt32BigEndian(data, offset);
            long chunkEnd = (long)offset + 12L + length;
            if (length < 0 || chunkEnd > data.Length) {
                unsupportedReason = "PNG chunk length is invalid.";
                return false;
            }

            string type = Encoding.ASCII.GetString(data, offset + 4, 4);
            int chunkData = offset + 8;
            if (!IsPngChunkCrcValid(data, offset, length)) {
                unsupportedReason = "PNG chunk CRC is invalid.";
                return false;
            }

            if (type == "IHDR") {
                if (length < 13) {
                    unsupportedReason = "PNG IHDR chunk is invalid.";
                    return false;
                }

                width = ReadInt32BigEndian(data, chunkData);
                height = ReadInt32BigEndian(data, chunkData + 4);
                bitDepth = data[chunkData + 8];
                colorType = data[chunkData + 9];
                compression = data[chunkData + 10];
                filter = data[chunkData + 11];
                interlace = data[chunkData + 12];
            } else if (type == "IDAT") {
                idat.Write(data, chunkData, length);
            } else if (type == "PLTE") {
                if (length == 0 || length % 3 != 0 || length > 768) {
                    unsupportedReason = "PNG palette chunk is invalid.";
                    return false;
                }

                palette = new byte[length];
                Buffer.BlockCopy(data, chunkData, palette, 0, length);
            } else if (type == "tRNS") {
                transparency = new byte[length];
                Buffer.BlockCopy(data, chunkData, transparency, 0, length);
            } else if (type == "IEND") {
                break;
            }

            offset = (int)chunkEnd;
        }

        if (width <= 0 || height <= 0) {
            unsupportedReason = "PNG dimensions are missing.";
            return false;
        }

        if (idat.Length == 0) {
            unsupportedReason = "PNG image data is missing.";
            return false;
        }

        if (colorType == 0) {
            if (bitDepth != 1 && bitDepth != 2 && bitDepth != 4 && bitDepth != 8 && bitDepth != 16) {
                unsupportedReason = "Grayscale PNG images must use 1, 2, 4, 8, or 16 bits per pixel.";
                return false;
            }
        } else if (colorType == 3) {
            if (bitDepth != 1 && bitDepth != 2 && bitDepth != 4 && bitDepth != 8) {
                unsupportedReason = "Indexed-color PNG images must use 1, 2, 4, or 8 bits per pixel.";
                return false;
            }
        } else if (bitDepth != 8 && bitDepth != 16) {
            unsupportedReason = "Only 8-bit and 16-bit PNG images are currently supported for grayscale-alpha, RGB, and RGBA PNG images.";
            return false;
        }

        if (compression != 0 || filter != 0) {
            unsupportedReason = "Unsupported PNG compression or filter method.";
            return false;
        }

        if (interlace != 0 && interlace != 1) {
            unsupportedReason = "Unsupported PNG interlace method.";
            return false;
        }

        if (transparency != null && (colorType == 4 || colorType == 6)) {
            unsupportedReason = "PNG transparency chunks are not valid for grayscale-alpha or RGBA PNG images.";
            return false;
        }

        if (!TryValidatePngResourceLimits(width, height, bitDepth, colorType, out unsupportedReason)) {
            return false;
        }

        byte[] streamData = idat.ToArray();
        if (interlace == 1 &&
            !TryNormalizeAdam7PngData(streamData, width, height, bitDepth, colorType, out streamData, out unsupportedReason)) {
            return false;
        }

        int colors;
        string colorSpace;
        if (colorType == 0) {
            if (bitDepth == 16) {
                return TryExpand16BitPng(streamData, width, height, colorType, transparency, out image, out unsupportedReason);
            }

            if (bitDepth != 8) {
                return TryExpandPackedGrayscalePng(streamData, width, height, bitDepth, transparency, out image, out unsupportedReason);
            }

            if (transparency != null) {
                return TrySplitPngTransparency(streamData, width, height, colorType, transparency, out image, out unsupportedReason);
            }

            colors = 1;
            colorSpace = "/DeviceGray";
        } else if (colorType == 2) {
            if (bitDepth == 16) {
                return TryExpand16BitPng(streamData, width, height, colorType, transparency, out image, out unsupportedReason);
            }

            if (transparency != null) {
                return TrySplitPngTransparency(streamData, width, height, colorType, transparency, out image, out unsupportedReason);
            }

            colors = 3;
            colorSpace = "/DeviceRGB";
        } else if (colorType == 3) {
            if (!TryExpandIndexedPng(streamData, width, height, bitDepth, palette, transparency, out image, out unsupportedReason)) {
                return false;
            }

            return true;
        } else if (colorType == 4 || colorType == 6) {
            if (bitDepth == 16) {
                return TryExpand16BitPng(streamData, width, height, colorType, transparency, out image, out unsupportedReason);
            }

            if (!TrySplitPngAlpha(streamData, width, height, colorType, out image, out unsupportedReason)) {
                return false;
            }

            return true;
        } else {
            unsupportedReason = "Only grayscale, grayscale-alpha, indexed-color, RGB, and RGBA PNG images are currently supported.";
            return false;
        }

        if (!TryValidatePngPassThroughData(streamData, width, height, colors, out unsupportedReason)) {
            return false;
        }

        image = new PdfImageStream {
            Data = streamData,
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = " /ColorSpace " + colorSpace +
                               " /BitsPerComponent 8 /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors " +
                               colors.ToString(CultureInfo.InvariantCulture) +
                               " /BitsPerComponent 8 /Columns " +
                               width.ToString(CultureInfo.InvariantCulture) +
                               " >>"
        };
        return true;
    }

    private static bool TryValidatePngPassThroughData(byte[] compressedData, int width, int height, int colors, out string? unsupportedReason) {
        if (!TryDecodePngData(compressedData, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, colors, includeFilterByte: true, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length != expectedLength) {
            unsupportedReason = "PNG image data length does not match the expected scanline size.";
            return false;
        }

        return true;
    }

    private static bool TrySplitPngTransparency(byte[] compressedData, int width, int height, int colorType, byte[] transparency, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        int sourceChannels = colorType == 0 ? 1 : 3;
        int requiredTransparencyLength = colorType == 0 ? 2 : 6;
        if (transparency.Length < requiredTransparencyLength) {
            unsupportedReason = colorType == 0
                ? "Grayscale PNG transparency chunk is invalid."
                : "RGB PNG transparency chunk is invalid.";
            return false;
        }

        if (!TryDecodePngData(compressedData, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, sourceChannels, includeFilterByte: true, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG image data ended before all transparency scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, width, height, sourceChannels, out var rawPixels, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, sourceChannels, includeFilterByte: true, out int baseRowsLength) ||
            !TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int alphaRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[baseRowsLength];
        byte[] alphaRows = new byte[alphaRowsLength];
        int transparentGray = ReadUInt16BigEndian(transparency, 0);
        int transparentRed = colorType == 2 ? ReadUInt16BigEndian(transparency, 0) : -1;
        int transparentGreen = colorType == 2 ? ReadUInt16BigEndian(transparency, 2) : -1;
        int transparentBlue = colorType == 2 ? ReadUInt16BigEndian(transparency, 4) : -1;

        for (int row = 0; row < height; row++) {
            int baseRowStart = row * (1 + width * sourceChannels);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            alphaRows[alphaRowStart] = 0;

            int sourceRowStart = row * width * sourceChannels;
            for (int pixel = 0; pixel < width; pixel++) {
                int sourcePixel = sourceRowStart + pixel * sourceChannels;
                int basePixel = baseRowStart + 1 + pixel * sourceChannels;
                Buffer.BlockCopy(rawPixels, sourcePixel, baseRows, basePixel, sourceChannels);

                bool isTransparent = colorType == 0
                    ? rawPixels[sourcePixel] == transparentGray
                    : rawPixels[sourcePixel] == transparentRed &&
                      rawPixels[sourcePixel + 1] == transparentGreen &&
                      rawPixels[sourcePixel + 2] == transparentBlue;
                alphaRows[alphaRowStart + 1 + pixel] = isTransparent ? (byte)0 : (byte)255;
            }
        }

        string colorSpace = colorType == 0 ? "/DeviceGray" : "/DeviceRGB";
        image = new PdfImageStream {
            Data = DeflateZlib(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix(colorSpace, sourceChannels, width),
            SoftMask = new PdfImageStream {
                Data = DeflateZlib(alphaRows),
                PixelWidth = width,
                PixelHeight = height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
            }
        };
        return true;
    }

    private static bool TryExpandPackedGrayscalePng(byte[] compressedData, int width, int height, int bitDepth, byte[]? transparency, out PdfImageStream image, out string? unsupportedReason) {
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

        if (!TryDecodePngData(compressedData, out byte[] decoded, out unsupportedReason)) {
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

        if (!TryUnfilterPngRows(decoded, packedRowBytes, height, 1, out var packedRows, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int grayscaleRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[grayscaleRowsLength];
        byte[]? alphaRows = transparency != null ? new byte[grayscaleRowsLength] : null;
        for (int row = 0; row < height; row++) {
            int baseRowStart = row * (1 + width);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * packedRowBytes;
            for (int pixel = 0; pixel < width; pixel++) {
                int sample = ReadPackedPngSample(packedRows, sourceRowStart, pixel, bitDepth);
                int targetOffset = baseRowStart + 1 + pixel;
                baseRows[targetOffset] = ScalePackedSampleToByte(sample, maxSample);
                if (alphaRows != null) {
                    alphaRows[alphaRowStart + 1 + pixel] = sample == transparentSample ? (byte)0 : (byte)255;
                }
            }
        }

        image = new PdfImageStream {
            Data = DeflateZlib(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
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

    private static bool TryExpandIndexedPng(byte[] compressedData, int width, int height, int bitDepth, byte[]? palette, byte[]? paletteAlpha, out PdfImageStream image, out string? unsupportedReason) {
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

        if (!TryDecodePngData(compressedData, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngRowByteCount(width, bitDepth, out int packedRowBytes) ||
            !TryGetPngScanlineLength(packedRowBytes, height, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG image data ended before all indexed-color scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, packedRowBytes, height, 1, out var packedRows, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, 3, includeFilterByte: true, out int rgbRowsLength) ||
            !TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int indexedAlphaRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[rgbRowsLength];
        byte[]? alphaRows = HasPaletteTransparency(paletteAlpha) ? new byte[indexedAlphaRowsLength] : null;
        for (int row = 0; row < height; row++) {
            int baseRowStart = row * (1 + width * 3);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            if (alphaRows != null) {
                alphaRows[alphaRowStart] = 0;
            }

            int sourceRowStart = row * packedRowBytes;
            for (int pixel = 0; pixel < width; pixel++) {
                int paletteIndex = ReadPackedPngSample(packedRows, sourceRowStart, pixel, bitDepth);
                if (paletteIndex >= paletteEntries) {
                    unsupportedReason = "Indexed-color PNG pixel references a palette entry that does not exist.";
                    return false;
                }

                int paletteOffset = paletteIndex * 3;
                int targetOffset = baseRowStart + 1 + pixel * 3;
                baseRows[targetOffset] = palette[paletteOffset];
                baseRows[targetOffset + 1] = palette[paletteOffset + 1];
                baseRows[targetOffset + 2] = palette[paletteOffset + 2];

                if (alphaRows != null) {
                    alphaRows[alphaRowStart + 1 + pixel] =
                        paletteAlpha != null && paletteIndex < paletteAlpha.Length ? paletteAlpha[paletteIndex] : (byte)255;
                }
            }
        }

        image = new PdfImageStream {
            Data = DeflateZlib(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceRGB", 3, width)
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

    private static bool HasPaletteTransparency(byte[]? paletteAlpha) =>
        paletteAlpha != null && paletteAlpha.Any(alpha => alpha < 255);

    private static int ReadPackedPngSample(byte[] packedRows, int rowStart, int pixelIndex, int bitDepth) {
        if (bitDepth == 8) {
            return packedRows[rowStart + pixelIndex];
        }

        int samplesPerByte = 8 / bitDepth;
        int sourceByte = packedRows[rowStart + pixelIndex / samplesPerByte];
        int shift = (samplesPerByte - 1 - (pixelIndex % samplesPerByte)) * bitDepth;
        int mask = (1 << bitDepth) - 1;
        return (sourceByte >> shift) & mask;
    }

    private static byte ScalePackedSampleToByte(int sample, int maxSample) =>
        (byte)Math.Round(sample * 255.0 / maxSample, MidpointRounding.AwayFromZero);

    private static int ReadUInt16BigEndian(byte[] data, int offset) =>
        offset + 2 <= data.Length
            ? (data[offset] << 8) | data[offset + 1]
            : 0;

    private static bool TrySplitPngAlpha(byte[] compressedData, int width, int height, int colorType, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        int sourceChannels = colorType == 4 ? 2 : 4;
        int baseChannels = colorType == 4 ? 1 : 3;
        if (!TryDecodePngData(compressedData, out byte[] decoded, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, sourceChannels, includeFilterByte: true, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG image data ended before all alpha scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, width, height, sourceChannels, out var rawPixels, out unsupportedReason)) {
            return false;
        }

        if (!TryGetPngCheckedLength(width, height, baseChannels, includeFilterByte: true, out int alphaBaseRowsLength) ||
            !TryGetPngCheckedLength(width, height, 1, includeFilterByte: true, out int splitAlphaRowsLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        byte[] baseRows = new byte[alphaBaseRowsLength];
        byte[] alphaRows = new byte[splitAlphaRowsLength];
        for (int row = 0; row < height; row++) {
            int baseRowStart = row * (1 + width * baseChannels);
            int alphaRowStart = row * (1 + width);
            baseRows[baseRowStart] = 0;
            alphaRows[alphaRowStart] = 0;

            int sourceRowStart = row * width * sourceChannels;
            for (int pixel = 0; pixel < width; pixel++) {
                int sourcePixel = sourceRowStart + pixel * sourceChannels;
                int basePixel = baseRowStart + 1 + pixel * baseChannels;
                int alphaPixel = alphaRowStart + 1 + pixel;

                for (int channel = 0; channel < baseChannels; channel++) {
                    baseRows[basePixel + channel] = rawPixels[sourcePixel + channel];
                }

                alphaRows[alphaPixel] = rawPixels[sourcePixel + sourceChannels - 1];
            }
        }

        string colorSpace = colorType == 4 ? "/DeviceGray" : "/DeviceRGB";
        image = new PdfImageStream {
            Data = DeflateZlib(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix(colorSpace, baseChannels, width),
            SoftMask = new PdfImageStream {
                Data = DeflateZlib(alphaRows),
                PixelWidth = width,
                PixelHeight = height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
            }
        };
        return true;
    }

    private static string BuildPngPredictorDictionarySuffix(string colorSpace, int colors, int width) =>
        " /ColorSpace " + colorSpace +
        " /BitsPerComponent 8 /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors " +
        colors.ToString(CultureInfo.InvariantCulture) +
        " /BitsPerComponent 8 /Columns " +
        width.ToString(CultureInfo.InvariantCulture) +
        " >>";

    private static bool TryUnfilterPngRows(byte[] decoded, int width, int height, int bytesPerPixel, out byte[] rawPixels, out string? unsupportedReason) {
        rawPixels = Array.Empty<byte>();
        unsupportedReason = null;

        if (!TryGetPngCheckedLength(width, height, bytesPerPixel, includeFilterByte: false, out int rawLength) ||
            !TryGetPngCheckedLength(width, height, bytesPerPixel, includeFilterByte: true, out int expectedLength)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        int stride = rawLength / height;
        int sourceRowLength = expectedLength / height;
        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG scanline data is incomplete.";
            return false;
        }

        rawPixels = new byte[rawLength];
        byte[] previous = new byte[stride];
        byte[] current = new byte[stride];

        for (int row = 0; row < height; row++) {
            int sourceRow = row * sourceRowLength;
            int filterType = decoded[sourceRow];
            Buffer.BlockCopy(decoded, sourceRow + 1, current, 0, stride);

            for (int i = 0; i < stride; i++) {
                int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
                int up = previous[i];
                int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;

                switch (filterType) {
                    case 0:
                        break;
                    case 1:
                        current[i] = unchecked((byte)(current[i] + left));
                        break;
                    case 2:
                        current[i] = unchecked((byte)(current[i] + up));
                        break;
                    case 3:
                        current[i] = unchecked((byte)(current[i] + ((left + up) >> 1)));
                        break;
                    case 4:
                        current[i] = unchecked((byte)(current[i] + PaethPredictor(left, up, upLeft)));
                        break;
                    default:
                        unsupportedReason = "Unsupported PNG scanline filter: " + filterType.ToString(CultureInfo.InvariantCulture) + ".";
                        return false;
                }
            }

            Buffer.BlockCopy(current, 0, rawPixels, row * stride, stride);
            Buffer.BlockCopy(current, 0, previous, 0, stride);
        }

        return true;
    }

    private static bool TryDecodePngData(byte[] compressedData, out byte[] decoded, out string? unsupportedReason) {
        if (!FlateDecoder.TryDecode(compressedData, MaxPngDecodedBytes, out decoded)) {
            unsupportedReason = "PNG image data exceeds the supported decompressed size limit.";
            return false;
        }

        unsupportedReason = null;
        return true;
    }

    private static bool TryValidatePngResourceLimits(int width, int height, int bitDepth, int colorType, out string? unsupportedReason) {
        unsupportedReason = null;
        long pixels = (long)width * height;
        if (pixels > MaxPngPixelCount) {
            unsupportedReason = "PNG dimensions exceed the supported pixel count limit.";
            return false;
        }

        if (!TryGetPngChannelCount(colorType, out int channels) ||
            !TryGetPngRowByteCount(width, channels * bitDepth, out int rowBytes) ||
            !TryGetPngScanlineLength(rowBytes, height, out int _)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        int expandedChannels = colorType == 3 || colorType == 6 ? 4 : Math.Max(channels, 1);
        if (colorType == 4) {
            expandedChannels = 2;
        }

        if (!TryGetPngCheckedLength(width, height, expandedChannels, includeFilterByte: true, out int _)) {
            unsupportedReason = "PNG dimensions exceed supported limits.";
            return false;
        }

        return true;
    }

    private static bool TryGetPngCheckedLength(int width, int height, int channels, bool includeFilterByte, out int length) {
        length = 0;
        long rowLength = ((long)width * channels) + (includeFilterByte ? 1 : 0);
        long totalLength = rowLength * height;
        if (rowLength > int.MaxValue || totalLength > int.MaxValue || totalLength > MaxPngExpandedBytes) {
            return false;
        }

        length = (int)totalLength;
        return true;
    }

    private static bool TryGetPngScanlineLength(int rowBytes, int height, out int length) {
        length = 0;
        long totalLength = ((long)rowBytes + 1L) * height;
        if (totalLength > int.MaxValue || totalLength > MaxPngExpandedBytes) {
            return false;
        }

        length = (int)totalLength;
        return true;
    }

    private static int PaethPredictor(int left, int up, int upLeft) {
        int p = left + up - upLeft;
        int pa = Math.Abs(p - left);
        int pb = Math.Abs(p - up);
        int pc = Math.Abs(p - upLeft);
        if (pa <= pb && pa <= pc) return left;
        return pb <= pc ? up : upLeft;
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static bool TryBuildImageStream(PageImage img, out PdfImageStream image, out string? unsupportedReason) {
        return TryBuildImageStream(img.Data, img.Info, img.W, img.H, out image, out unsupportedReason);
    }

    private static string BuildImageXObjectCacheKey(PdfImageStream image) {
        using var hash = System.Security.Cryptography.SHA256.Create();
        AppendImageStreamHash(hash, image);
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return ToHex(hash.Hash ?? Array.Empty<byte>());
    }

    private static void AppendImageStreamHash(System.Security.Cryptography.HashAlgorithm hash, PdfImageStream image) {
        AppendHashInt(hash, image.PixelWidth);
        AppendHashInt(hash, image.PixelHeight);
        AppendHashString(hash, image.DictionarySuffix);
        AppendHashBytes(hash, image.Data);
        if (image.SoftMask == null) {
            AppendHashByte(hash, 0);
            return;
        }

        AppendHashByte(hash, 1);
        AppendImageStreamHash(hash, image.SoftMask);
    }

    private static void AppendHashString(System.Security.Cryptography.HashAlgorithm hash, string value) =>
        AppendHashBytes(hash, Encoding.UTF8.GetBytes(value ?? string.Empty));

    private static void AppendHashByte(System.Security.Cryptography.HashAlgorithm hash, byte value) =>
        AppendHashBytes(hash, new[] { value });

    private static void AppendHashInt(System.Security.Cryptography.HashAlgorithm hash, int value) {
        byte[] bytes = new byte[] {
            (byte)((value >> 24) & 0xFF),
            (byte)((value >> 16) & 0xFF),
            (byte)((value >> 8) & 0xFF),
            (byte)(value & 0xFF)
        };
        AppendHashBytes(hash, bytes);
    }

    private static void AppendHashBytes(System.Security.Cryptography.HashAlgorithm hash, byte[] data) {
        AppendHashLength(hash, data.Length);
        if (data.Length > 0) {
            hash.TransformBlock(data, 0, data.Length, data, 0);
        }
    }

    private static void AppendHashLength(System.Security.Cryptography.HashAlgorithm hash, int length) {
        byte[] bytes = new byte[] {
            (byte)((length >> 24) & 0xFF),
            (byte)((length >> 16) & 0xFF),
            (byte)((length >> 8) & 0xFF),
            (byte)(length & 0xFF)
        };
        hash.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
    }

    private static string ToHex(byte[] bytes) {
        char[] chars = new char[bytes.Length * 2];
        const string hex = "0123456789abcdef";
        for (int i = 0; i < bytes.Length; i++) {
            chars[i * 2] = hex[bytes[i] >> 4];
            chars[i * 2 + 1] = hex[bytes[i] & 0xF];
        }

        return new string(chars);
    }

    internal static bool TryBuildImageStream(byte[] data, OfficeImageInfo info, double fallbackWidth, double fallbackHeight, out PdfImageStream image, out string? unsupportedReason) {
        unsupportedReason = null;

        if (info.Format == OfficeImageFormat.Png) {
            return TryGetPngImageData(data, out image, out unsupportedReason);
        }

        int pixelWidth = info.Width > 0 ? info.Width : Math.Max(1, (int)Math.Round(fallbackWidth));
        int pixelHeight = info.Height > 0 ? info.Height : Math.Max(1, (int)Math.Round(fallbackHeight));
        image = new PdfImageStream {
            Data = data,
            PixelWidth = pixelWidth,
            PixelHeight = pixelHeight,
            DictionarySuffix = " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode"
        };
        return true;
    }

    internal static string BuildImageObjectDictionary(PdfImageStream image, int? softMaskObjectId = null) {
        return PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(image, softMaskObjectId);
    }

    internal static PdfStream BuildImageXObject(PdfImageStream image, int? softMaskObjectNumber = null) {
        return PdfImageXObjectDictionaryBuilder.BuildStreamObject(image, softMaskObjectNumber);
    }

    private static bool IsPngChunkCrcValid(byte[] data, int chunkOffset, int chunkLength) {
        long crcOffsetLong = (long)chunkOffset + 8L + chunkLength;
        if (crcOffsetLong < 0 || crcOffsetLong + 4L > data.Length) {
            return false;
        }

        int crcOffset = (int)crcOffsetLong;
        uint expectedCrc = ReadUInt32BigEndian(data, crcOffset);
        uint actualCrc = Crc32(data, chunkOffset + 4, chunkLength + 4);
        return expectedCrc == actualCrc;
    }

    private static uint Crc32(byte[] data, int offset, int length) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < length; i++) {
            crc ^= data[offset + i];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) == 1 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
            }
        }

        return ~crc;
    }

    private static uint ReadUInt32BigEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3]
            : 0;

    private static bool IsPng(byte[] data) =>
        data.Length >= 8 &&
        data[0] == 137 &&
        data[1] == 80 &&
        data[2] == 78 &&
        data[3] == 71 &&
        data[4] == 13 &&
        data[5] == 10 &&
        data[6] == 26 &&
        data[7] == 10;

    private static int ReadInt32BigEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]
            : 0;
}
