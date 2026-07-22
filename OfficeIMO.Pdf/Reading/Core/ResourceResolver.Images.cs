using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    /// <summary>Determines whether the managed image projection can normalize an authored image color space.</summary>
    internal static bool CanProjectImageColorSpace(
        PdfDictionary image,
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects) {
        if (image.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) &&
            ResolveObject(imageMaskObject, objects) is PdfBoolean { Value: true }) {
            return true;
        }

        PdfObject? authoredColorSpace = image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
            ? colorSpaceObject
            : null;
        PdfObject? effectiveColorSpace = ResolveColorSpaceResource(authoredColorSpace, resources, objects);
        int bitsPerComponent = (int)(image.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        if (PdfIndexedImageNormalizer.CanNormalizeColorSpace(effectiveColorSpace, bitsPerComponent, objects)) {
            return true;
        }

        string colorSpaceName = GetNameOrEmpty(effectiveColorSpace, objects);
        return PdfImageColorSpaceNormalization.TryResolve(
            effectiveColorSpace,
            colorSpaceName,
            objects,
            out _);
    }

    private static bool TryBuildExtractedImageMaskPng(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects,
        OfficeColor? imageMaskColor,
        out byte[] pngBytes) {
        if (imageMaskColor.HasValue) {
            return TryBuildPngFileFromImageMask(stream, width, height, bitsPerComponent, objects, imageMaskColor.Value, out pngBytes);
        }

        return PdfImageMaskNormalizer.TryBuildPngFile(width, height, stream, objects, out pngBytes);
    }

    private static bool TryBuildPngFileFromImageMask(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects,
        OfficeColor imageMaskColor,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (bitsPerComponent is not (0 or 1)) {
            return false;
        }
        if (!PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, 4, out int pixelCount, out int scanlineBytes)) {
            return false;
        }

        byte[] maskPixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        byte[] scanlines = new byte[scanlineBytes];
        for (int sampleIndex = 0; sampleIndex < pixelCount; sampleIndex++) {
            if (!TryReadIndexedSample(maskPixels, width, sampleIndex, 1, out int sample)) {
                return false;
            }

            int row = sampleIndex / width;
            int column = sampleIndex - row * width;
            int outputPixel = row * (1 + width * 4) + 1 + column * 4;
            scanlines[row * (1 + width * 4)] = 0;
            scanlines[outputPixel] = imageMaskColor.R;
            scanlines[outputPixel + 1] = imageMaskColor.G;
            scanlines[outputPixel + 2] = imageMaskColor.B;
            scanlines[outputPixel + 3] = DecodeImageMaskAlpha(stream.Dictionary, sample, objects);
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            8,
            6,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryBuildPngFileFromDeviceColor(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        int colorType,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (bitsPerComponent != 8) {
            return false;
        }

        int baseColors;
        int outputColorType;
        bool hasSoftMask = stream.Dictionary.Items.ContainsKey("SMask");
        PdfColorKeyMask? colorKeyMask = TryReadColorKeyMask(stream.Dictionary, colorType == 0 ? 1 : 3, objects, out PdfColorKeyMask parsedColorKeyMask)
            ? parsedColorKeyMask
            : null;
        bool hasAlpha = hasSoftMask || colorKeyMask.HasValue;
        if (colorType == 0) {
            baseColors = 1;
            outputColorType = hasAlpha ? 4 : 0;
        } else if (colorType == 2) {
            baseColors = 3;
            outputColorType = hasAlpha ? 6 : 2;
        } else {
            return false;
        }

        byte[] basePixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        int expectedBaseLength = width * height * baseColors;
        if (basePixels.Length < expectedBaseLength) {
            return false;
        }

        byte[]? alphaPixels = null;
        if (hasSoftMask &&
            !TryDecodeSoftMask(stream, width, height, bitsPerComponent, objects, out alphaPixels)) {
            return false;
        }

        int outputChannels = hasAlpha ? baseColors + 1 : baseColors;
        byte[] scanlines = new byte[(1 + width * outputChannels) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + width * outputChannels);
            int baseRow = row * width * baseColors;
            int alphaRow = row * width;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int basePixel = baseRow + pixel * baseColors;
                int outputPixel = outputRow + 1 + pixel * outputChannels;
                for (int channel = 0; channel < baseColors; channel++) {
                    scanlines[outputPixel + channel] = DecodeImageComponent(stream.Dictionary, channel, basePixels[basePixel + channel], objects);
                }

                if (hasAlpha) {
                    if (alphaPixels is not null && alphaPixels.Length <= alphaRow + pixel) {
                        return false;
                    }

                    scanlines[outputPixel + baseColors] = ResolveImageAlpha(basePixels, basePixel, baseColors, colorKeyMask, alphaPixels, alphaRow + pixel);
                }
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            8,
            outputColorType,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryBuildPngFileFromIndexed(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        PdfObject? colorSpaceObject,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (bitsPerComponent is not (1 or 2 or 4 or 8) ||
            !TryReadIndexedColorSpace(colorSpaceObject, objects, out var palette, out int highValue)) {
            return false;
        }

        byte[] indexedPixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        int expectedSampleCount = width * height;
        bool hasSoftMask = stream.Dictionary.Items.ContainsKey("SMask");
        byte[]? alphaPixels = null;
        if (hasSoftMask &&
            !TryDecodeSoftMask(stream, width, height, 8, objects, out alphaPixels)) {
            return false;
        }

        PdfColorKeyMask? colorKeyMask = TryReadColorKeyMask(stream.Dictionary, 1, objects, out PdfColorKeyMask parsedColorKeyMask)
            ? parsedColorKeyMask
            : null;
        bool hasAlpha = hasSoftMask || colorKeyMask.HasValue;
        int outputChannels = hasAlpha ? 4 : 3;
        int colorType = hasAlpha ? 6 : 2;
        byte[] scanlines = new byte[(1 + width * outputChannels) * height];
        for (int sampleIndex = 0; sampleIndex < expectedSampleCount; sampleIndex++) {
            if (!TryReadIndexedSample(indexedPixels, width, sampleIndex, bitsPerComponent, out int sample)) {
                return false;
            }

            int paletteIndex = MapIndexedSample(stream.Dictionary, sample, bitsPerComponent, highValue, objects);
            int paletteOffset = paletteIndex * 3;
            int row = sampleIndex / width;
            int column = sampleIndex - row * width;
            int outputPixel = row * (1 + width * outputChannels) + 1 + column * outputChannels;
            scanlines[row * (1 + width * outputChannels)] = 0;
            scanlines[outputPixel] = palette[paletteOffset];
            scanlines[outputPixel + 1] = palette[paletteOffset + 1];
            scanlines[outputPixel + 2] = palette[paletteOffset + 2];
            if (hasAlpha) {
                if (alphaPixels is not null && alphaPixels.Length <= sampleIndex) {
                    return false;
                }

                scanlines[outputPixel + 3] = ResolveImageAlpha(paletteIndex, colorKeyMask, alphaPixels, sampleIndex);
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            8,
            colorType,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryReadIndexedColorSpace(
        PdfObject? colorSpaceObject,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] rgbPalette,
        out int highValue) {
        rgbPalette = Array.Empty<byte>();
        highValue = 0;
        if (ResolveObject(colorSpaceObject, objects) is not PdfArray colorSpace ||
            colorSpace.Items.Count < 4 ||
            ResolveObject(colorSpace.Items[2], objects) is not PdfNumber highValueNumber ||
            !TryReadIndexedBaseColorSpace(colorSpace.Items[1], objects, out var baseColorSpace, out int baseComponentCount) ||
            !TryReadIndexedLookup(colorSpace.Items[3], objects, out var lookupBytes)) {
            return false;
        }

        highValue = (int)highValueNumber.Value;
        if (highValue < 0 || highValue > 255) {
            return false;
        }

        int paletteEntries = highValue + 1;
        if (lookupBytes.Length < paletteEntries * baseComponentCount) {
            return false;
        }

        rgbPalette = new byte[paletteEntries * 3];
        for (int index = 0; index < paletteEntries; index++) {
            int source = index * baseComponentCount;
            int target = index * 3;
            switch (baseColorSpace) {
                case IndexedBaseColorSpace.DeviceRgb:
                    rgbPalette[target] = lookupBytes[source];
                    rgbPalette[target + 1] = lookupBytes[source + 1];
                    rgbPalette[target + 2] = lookupBytes[source + 2];
                    break;
                case IndexedBaseColorSpace.DeviceGray:
                    rgbPalette[target] = lookupBytes[source];
                    rgbPalette[target + 1] = lookupBytes[source];
                    rgbPalette[target + 2] = lookupBytes[source];
                    break;
                case IndexedBaseColorSpace.DeviceCmyk:
                    byte black = lookupBytes[source + 3];
                    rgbPalette[target] = ConvertCmykComponent(lookupBytes[source], black);
                    rgbPalette[target + 1] = ConvertCmykComponent(lookupBytes[source + 1], black);
                    rgbPalette[target + 2] = ConvertCmykComponent(lookupBytes[source + 2], black);
                    break;
                default:
                    return false;
            }
        }

        return true;
    }

    private static bool TryReadIndexedBaseColorSpace(
        PdfObject? baseColorSpaceObject,
        Dictionary<int, PdfIndirectObject> objects,
        out IndexedBaseColorSpace colorSpace,
        out int componentCount) {
        colorSpace = IndexedBaseColorSpace.Unsupported;
        componentCount = 0;
        string name = GetNameOrEmpty(baseColorSpaceObject, objects);
        switch (name) {
            case "DeviceRGB":
            case "RGB":
                colorSpace = IndexedBaseColorSpace.DeviceRgb;
                componentCount = 3;
                return true;
            case "DeviceGray":
            case "G":
                colorSpace = IndexedBaseColorSpace.DeviceGray;
                componentCount = 1;
                return true;
            case "DeviceCMYK":
            case "CMYK":
                colorSpace = IndexedBaseColorSpace.DeviceCmyk;
                componentCount = 4;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadIndexedLookup(
        PdfObject? lookupObject,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] lookupBytes) {
        lookupBytes = Array.Empty<byte>();
        switch (ResolveObject(lookupObject, objects)) {
            case PdfStringObj lookupString:
                lookupBytes = lookupString.RawBytes;
                return lookupBytes.Length > 0;
            case PdfStream lookupStream:
                lookupBytes = Filters.StreamDecoder.Decode(lookupStream.Dictionary, lookupStream.Data, objects);
                return lookupBytes.Length > 0;
            default:
                return false;
        }
    }

    private static bool TryReadIndexedSample(byte[] pixels, int width, int sampleIndex, int bitsPerComponent, out int sample) {
        sample = 0;
        int row = sampleIndex / width;
        int column = sampleIndex - row * width;
        if (bitsPerComponent == 8) {
            if (sampleIndex >= pixels.Length) {
                return false;
            }

            sample = pixels[sampleIndex];
            return true;
        }

        int rowBitCount = width * bitsPerComponent;
        int rowByteCount = (rowBitCount + 7) / 8;
        int bitOffset = row * rowByteCount * 8 + column * bitsPerComponent;
        int byteIndex = bitOffset / 8;
        if (byteIndex >= pixels.Length) {
            return false;
        }

        int shift = 8 - bitsPerComponent - bitOffset % 8;
        int mask = (1 << bitsPerComponent) - 1;
        sample = pixels[byteIndex] >> shift & mask;
        return true;
    }

    private static int MapIndexedSample(
        PdfDictionary imageDictionary,
        int sample,
        int bitsPerComponent,
        int highValue,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!imageDictionary.Items.TryGetValue("Decode", out var decodeObject) ||
            ResolveObject(decodeObject, objects) is not PdfArray decodeArray ||
            decodeArray.Items.Count < 2 ||
            ResolveObject(decodeArray.Items[0], objects) is not PdfNumber decodeMin ||
            ResolveObject(decodeArray.Items[1], objects) is not PdfNumber decodeMax) {
            return Clamp(sample, 0, highValue);
        }

        int maxSample = (1 << bitsPerComponent) - 1;
        double mapped = decodeMin.Value + sample * (decodeMax.Value - decodeMin.Value) / maxSample;
        return Clamp((int)Math.Round(mapped), 0, highValue);
    }

    private static byte DecodeImageComponent(
        PdfDictionary imageDictionary,
        int channel,
        byte sample,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!imageDictionary.Items.TryGetValue("Decode", out var decodeObject) ||
            ResolveObject(decodeObject, objects) is not PdfArray decodeArray) {
            return sample;
        }

        int minimumIndex = channel * 2;
        int maximumIndex = minimumIndex + 1;
        if (decodeArray.Items.Count <= maximumIndex ||
            ResolveObject(decodeArray.Items[minimumIndex], objects) is not PdfNumber decodeMin ||
            ResolveObject(decodeArray.Items[maximumIndex], objects) is not PdfNumber decodeMax) {
            return sample;
        }

        double mapped = decodeMin.Value + sample / 255D * (decodeMax.Value - decodeMin.Value);
        return ToByte(mapped);
    }

    private static byte DecodeImageMaskAlpha(
        PdfDictionary imageDictionary,
        int sample,
        Dictionary<int, PdfIndirectObject> objects) {
        double decodeMin = 0D;
        double decodeMax = 1D;
        if (imageDictionary.Items.TryGetValue("Decode", out var decodeObject) &&
            ResolveObject(decodeObject, objects) is PdfArray decodeArray &&
            decodeArray.Items.Count >= 2 &&
            ResolveObject(decodeArray.Items[0], objects) is PdfNumber decodeMinNumber &&
            ResolveObject(decodeArray.Items[1], objects) is PdfNumber decodeMaxNumber) {
            decodeMin = decodeMinNumber.Value;
            decodeMax = decodeMaxNumber.Value;
        }

        double mapped = decodeMin + sample * (decodeMax - decodeMin);
        return mapped >= 0.5D ? (byte)255 : (byte)0;
    }

    private static byte ResolveImageAlpha(byte[] samples, int sampleOffset, int componentCount, PdfColorKeyMask? colorKeyMask, byte[]? alphaPixels, int alphaIndex) {
        if (colorKeyMask.HasValue && colorKeyMask.Value.Matches(samples, sampleOffset, componentCount)) {
            return 0;
        }

        return alphaPixels is null ? (byte)255 : alphaPixels[alphaIndex];
    }

    private static byte ResolveImageAlpha(int sample, PdfColorKeyMask? colorKeyMask, byte[]? alphaPixels, int alphaIndex) {
        if (colorKeyMask.HasValue && colorKeyMask.Value.Matches(sample)) {
            return 0;
        }

        return alphaPixels is null ? (byte)255 : alphaPixels[alphaIndex];
    }

    private static bool TryReadColorKeyMask(
        PdfDictionary imageDictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfColorKeyMask colorKeyMask) {
        colorKeyMask = default;
        if (componentCount <= 0 ||
            !imageDictionary.Items.TryGetValue("Mask", out PdfObject? maskObject) ||
            ResolveObject(maskObject, objects) is not PdfArray maskArray ||
            maskArray.Items.Count < componentCount * 2) {
            return false;
        }

        var ranges = new int[componentCount * 2];
        for (int i = 0; i < ranges.Length; i++) {
            if (ResolveObject(maskArray.Items[i], objects) is not PdfNumber number) {
                return false;
            }

            ranges[i] = Clamp((int)Math.Round(number.Value), 0, 255);
        }

        colorKeyMask = new PdfColorKeyMask(ranges);
        return true;
    }

    private static byte ToByte(double value) {
        if (value < 0D) {
            value = 0D;
        } else if (value > 1D) {
            value = 1D;
        }

        return (byte)Math.Round(value * 255D);
    }

    private static int Clamp(int value, int min, int max) {
        if (value < min) {
            return min;
        }

        return value > max ? max : value;
    }

    private enum IndexedBaseColorSpace {
        Unsupported,
        DeviceRgb,
        DeviceGray,
        DeviceCmyk
    }

    private readonly struct PdfColorKeyMask {
        private readonly int[] _ranges;

        public PdfColorKeyMask(int[] ranges) {
            _ranges = ranges;
        }

        public bool Matches(byte[] samples, int sampleOffset, int componentCount) {
            if (_ranges.Length < componentCount * 2 || samples.Length < sampleOffset + componentCount) {
                return false;
            }

            for (int component = 0; component < componentCount; component++) {
                int sample = samples[sampleOffset + component];
                int rangeOffset = component * 2;
                if (sample < _ranges[rangeOffset] || sample > _ranges[rangeOffset + 1]) {
                    return false;
                }
            }

            return true;
        }

        public bool Matches(int sample) =>
            _ranges.Length >= 2 &&
            sample >= _ranges[0] &&
            sample <= _ranges[1];
    }
}
