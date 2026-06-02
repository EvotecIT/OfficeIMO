using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
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
        var idat = new MemoryStream();

        while (offset + 12 <= data.Length) {
            int length = ReadInt32BigEndian(data, offset);
            if (length < 0 || offset + 12 + length > data.Length) {
                unsupportedReason = "PNG chunk length is invalid.";
                return false;
            }

            string type = Encoding.ASCII.GetString(data, offset + 4, 4);
            int chunkData = offset + 8;
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
            } else if (type == "IEND") {
                break;
            }

            offset += 12 + length;
        }

        if (width <= 0 || height <= 0) {
            unsupportedReason = "PNG dimensions are missing.";
            return false;
        }

        if (idat.Length == 0) {
            unsupportedReason = "PNG image data is missing.";
            return false;
        }

        if (bitDepth != 8) {
            unsupportedReason = "Only 8-bit PNG images are currently supported.";
            return false;
        }

        if (compression != 0 || filter != 0) {
            unsupportedReason = "Unsupported PNG compression or filter method.";
            return false;
        }

        if (interlace != 0) {
            unsupportedReason = "Interlaced PNG images are not currently supported.";
            return false;
        }

        int colors;
        string colorSpace;
        if (colorType == 0) {
            colors = 1;
            colorSpace = "/DeviceGray";
        } else if (colorType == 2) {
            colors = 3;
            colorSpace = "/DeviceRGB";
        } else if (colorType == 4 || colorType == 6) {
            if (!TrySplitPngAlpha(idat.ToArray(), width, height, colorType, out image, out unsupportedReason)) {
                return false;
            }

            return true;
        } else {
            unsupportedReason = "Only grayscale, grayscale-alpha, RGB, and RGBA PNG images are currently supported.";
            return false;
        }

        byte[] streamData = idat.ToArray();
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

    private static bool TrySplitPngAlpha(byte[] compressedData, int width, int height, int colorType, out PdfImageStream image, out string? unsupportedReason) {
        image = new PdfImageStream();
        unsupportedReason = null;

        int sourceChannels = colorType == 4 ? 2 : 4;
        int baseChannels = colorType == 4 ? 1 : 3;
        byte[] decoded = FlateDecoder.Decode(compressedData);
        int expectedRowLength = 1 + width * sourceChannels;
        int expectedLength = expectedRowLength * height;
        if (decoded.Length < expectedLength) {
            unsupportedReason = "PNG image data ended before all alpha scanlines were decoded.";
            return false;
        }

        if (!TryUnfilterPngRows(decoded, width, height, sourceChannels, out var rawPixels, out unsupportedReason)) {
            return false;
        }

        byte[] baseRows = new byte[(1 + width * baseChannels) * height];
        byte[] alphaRows = new byte[(1 + width) * height];
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
            Data = DeflateZlibStored(baseRows),
            PixelWidth = width,
            PixelHeight = height,
            DictionarySuffix = BuildPngPredictorDictionarySuffix(colorSpace, baseChannels, width),
            SoftMask = new PdfImageStream {
                Data = DeflateZlibStored(alphaRows),
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

        int stride = width * bytesPerPixel;
        int sourceRowLength = stride + 1;
        if (decoded.Length < sourceRowLength * height) {
            unsupportedReason = "PNG scanline data is incomplete.";
            return false;
        }

        rawPixels = new byte[stride * height];
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

    private static int PaethPredictor(int left, int up, int upLeft) {
        int p = left + up - upLeft;
        int pa = Math.Abs(p - left);
        int pb = Math.Abs(p - up);
        int pc = Math.Abs(p - upLeft);
        if (pa <= pb && pa <= pc) return left;
        return pb <= pc ? up : upLeft;
    }

    private static byte[] DeflateZlibStored(byte[] data) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);

        int offset = 0;
        do {
            int blockLength = Math.Min(65535, data.Length - offset);
            bool final = offset + blockLength >= data.Length;
            ms.WriteByte(final ? (byte)1 : (byte)0);
            ms.WriteByte((byte)(blockLength & 0xFF));
            ms.WriteByte((byte)((blockLength >> 8) & 0xFF));
            int nlen = blockLength ^ 0xFFFF;
            ms.WriteByte((byte)(nlen & 0xFF));
            ms.WriteByte((byte)((nlen >> 8) & 0xFF));
            ms.Write(data, offset, blockLength);
            offset += blockLength;
        } while (offset < data.Length);

        uint adler = Adler32(data);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
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
