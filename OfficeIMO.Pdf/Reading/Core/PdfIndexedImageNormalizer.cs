using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfIndexedImageNormalizer {
    internal static bool CanNormalizeColorSpace(
        PdfObject? colorSpaceObj,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects) =>
        TryResolveIndexedPalette(colorSpaceObj, bitsPerComponent, objects, out _);

    internal static bool TryBuildPngFile(
        PdfObject? colorSpaceObj,
        int width,
        int height,
        int bitsPerComponent,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (width <= 0 ||
            height <= 0 ||
            !TryResolveIndexedPalette(colorSpaceObj, bitsPerComponent, objects, out var indexedPalette)) {
            return false;
        }

        if (!TryReadDecodedStreamBytes(stream, objects, out var indexedPixels)) {
            return false;
        }

        var decodeTransform = PdfImageDecodeTransform.CreateIndexed(stream.Dictionary, indexedPalette.Length / 3 - 1, objects);
        if (stream.Dictionary.Items.ContainsKey("SMask")) {
            return TryBuildPngFileFromIndexedPixelsWithSoftMask(width, height, bitsPerComponent, indexedPalette, decodeTransform, indexedPixels, stream, objects, out pngBytes);
        }

        var colorKeyMask = PdfImageColorKeyMask.Create(stream.Dictionary, 1, objects);
        if (colorKeyMask is not null) {
            return TryBuildPngFileFromIndexedPixelsWithColorKeyMask(width, height, bitsPerComponent, indexedPalette, decodeTransform, colorKeyMask, indexedPixels, out pngBytes);
        }

        return TryBuildPngFileFromIndexedPixels(width, height, bitsPerComponent, indexedPalette, decodeTransform, indexedPixels, out pngBytes);
    }

    private static bool TryResolveIndexedPalette(
        PdfObject? colorSpaceObj,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] rgbPalette) {
        rgbPalette = Array.Empty<byte>();
        if (bitsPerComponent != 1 && bitsPerComponent != 2 && bitsPerComponent != 4 && bitsPerComponent != 8) {
            return false;
        }

        if (ResolveObject(colorSpaceObj, objects) is not PdfArray colorSpaceArray ||
            colorSpaceArray.Items.Count < 4 ||
            ResolveObject(colorSpaceArray.Items[0], objects) is not PdfName indexedName ||
            (!string.Equals(indexedName.Name, "Indexed", StringComparison.Ordinal) &&
             !string.Equals(indexedName.Name, "I", StringComparison.Ordinal)) ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfName baseColorSpace ||
            ResolveObject(colorSpaceArray.Items[2], objects) is not PdfNumber highValueNumber) {
            return false;
        }

        int highValue = (int)highValueNumber.Value;
        if (highValue < 0 || highValue > 255) {
            return false;
        }

        int baseComponentCount;
        if (string.Equals(baseColorSpace.Name, "DeviceGray", StringComparison.Ordinal)) {
            baseComponentCount = 1;
        } else if (string.Equals(baseColorSpace.Name, "DeviceRGB", StringComparison.Ordinal)) {
            baseComponentCount = 3;
        } else if (string.Equals(baseColorSpace.Name, "DeviceCMYK", StringComparison.Ordinal)) {
            baseComponentCount = 4;
        } else {
            return false;
        }

        if (!TryReadIndexedLookupBytes(colorSpaceArray.Items[3], objects, out var lookupBytes)) {
            return false;
        }

        int paletteEntryCount = highValue + 1;
        int expectedLookupLength = paletteEntryCount * baseComponentCount;
        if (lookupBytes.Length < expectedLookupLength) {
            return false;
        }

        rgbPalette = new byte[paletteEntryCount * 3];
        for (int entry = 0; entry < paletteEntryCount; entry++) {
            int lookupOffset = entry * baseComponentCount;
            int paletteOffset = entry * 3;
            if (baseComponentCount == 1) {
                byte gray = lookupBytes[lookupOffset];
                rgbPalette[paletteOffset] = gray;
                rgbPalette[paletteOffset + 1] = gray;
                rgbPalette[paletteOffset + 2] = gray;
            } else if (baseComponentCount == 3) {
                rgbPalette[paletteOffset] = lookupBytes[lookupOffset];
                rgbPalette[paletteOffset + 1] = lookupBytes[lookupOffset + 1];
                rgbPalette[paletteOffset + 2] = lookupBytes[lookupOffset + 2];
            } else {
                byte c = lookupBytes[lookupOffset];
                byte m = lookupBytes[lookupOffset + 1];
                byte y = lookupBytes[lookupOffset + 2];
                byte k = lookupBytes[lookupOffset + 3];
                rgbPalette[paletteOffset] = ConvertDeviceCmykComponentToRgb(c, k);
                rgbPalette[paletteOffset + 1] = ConvertDeviceCmykComponentToRgb(m, k);
                rgbPalette[paletteOffset + 2] = ConvertDeviceCmykComponentToRgb(y, k);
            }
        }

        return true;
    }

    private static bool TryReadIndexedLookupBytes(
        PdfObject? lookupObject,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] lookupBytes) {
        lookupBytes = Array.Empty<byte>();
        PdfObject? resolvedLookup = ResolveObject(lookupObject, objects);
        if (resolvedLookup is PdfStringObj lookupString) {
            lookupBytes = lookupString.RawBytes;
            return true;
        }

        if (resolvedLookup is PdfStream lookupStream) {
            if (Filters.StreamDecoder.GetUnsupportedFilters(lookupStream.Dictionary, objects).Count != 0) {
                return false;
            }

            lookupBytes = Filters.StreamDecoder.Decode(lookupStream.Dictionary, lookupStream.Data, objects);
            return lookupBytes.Length > 0;
        }

        return false;
    }

    private static bool TryReadDecodedStreamBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] bytes) {
        bytes = Array.Empty<byte>();
        if (stream.Dictionary.Items.TryGetValue("Filter", out _)) {
            if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) {
                return false;
            }

            bytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        } else {
            bytes = stream.Data;
        }

        return bytes.Length > 0;
    }

    private static bool TryBuildPngFileFromIndexedPixels(
        int width,
        int height,
        int bitsPerComponent,
        byte[] rgbPalette,
        PdfImageDecodeTransform? decodeTransform,
        byte[] indexedPixels,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (indexedPixels.Length == 0 || rgbPalette.Length == 0 || rgbPalette.Length % 3 != 0) {
            return false;
        }

        long sourceRowLengthLong = ((long)width * bitsPerComponent + 7) / 8;
        long expectedLengthLong = sourceRowLengthLong * height;
        long outputRowLengthLong = (long)width * 3;
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (indexedPixels.Length < expectedLength) {
            return false;
        }

        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        int paletteEntryCount = rgbPalette.Length / 3;
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int paletteIndex = ReadIndexedPixel(indexedPixels, sourceRow, pixel, bitsPerComponent);
                if (decodeTransform is not null) {
                    paletteIndex = decodeTransform.TransformIndexedSample(paletteIndex, bitsPerComponent, paletteEntryCount - 1);
                }

                if (paletteIndex < 0 || paletteIndex >= paletteEntryCount) {
                    return false;
                }

                int paletteOffset = paletteIndex * 3;
                int outputPixel = outputRow + 1 + pixel * 3;
                scanlines[outputPixel] = rgbPalette[paletteOffset];
                scanlines[outputPixel + 1] = rgbPalette[paletteOffset + 1];
                scanlines[outputPixel + 2] = rgbPalette[paletteOffset + 2];
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            8,
            2,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryBuildPngFileFromIndexedPixelsWithColorKeyMask(
        int width,
        int height,
        int bitsPerComponent,
        byte[] rgbPalette,
        PdfImageDecodeTransform? decodeTransform,
        PdfImageColorKeyMask colorKeyMask,
        byte[] indexedPixels,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (indexedPixels.Length == 0 || rgbPalette.Length == 0 || rgbPalette.Length % 3 != 0) {
            return false;
        }

        long sourceRowLengthLong = ((long)width * bitsPerComponent + 7) / 8;
        long expectedLengthLong = sourceRowLengthLong * height;
        long outputRowLengthLong = (long)width * 4;
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (indexedPixels.Length < expectedLength) {
            return false;
        }

        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        int paletteEntryCount = rgbPalette.Length / 3;
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int rawPaletteIndex = ReadIndexedPixel(indexedPixels, sourceRow, pixel, bitsPerComponent);
                int paletteIndex = rawPaletteIndex;
                if (decodeTransform is not null) {
                    paletteIndex = decodeTransform.TransformIndexedSample(paletteIndex, bitsPerComponent, paletteEntryCount - 1);
                }

                if (paletteIndex < 0 || paletteIndex >= paletteEntryCount) {
                    return false;
                }

                int paletteOffset = paletteIndex * 3;
                int outputPixel = outputRow + 1 + pixel * 4;
                scanlines[outputPixel] = rgbPalette[paletteOffset];
                scanlines[outputPixel + 1] = rgbPalette[paletteOffset + 1];
                scanlines[outputPixel + 2] = rgbPalette[paletteOffset + 2];
                scanlines[outputPixel + 3] = colorKeyMask.IsTransparentSample(rawPaletteIndex)
                    ? (byte)0
                    : (byte)255;
            }
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

    private static bool TryBuildPngFileFromIndexedPixelsWithSoftMask(
        int width,
        int height,
        int bitsPerComponent,
        byte[] rgbPalette,
        PdfImageDecodeTransform? decodeTransform,
        byte[] indexedPixels,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (!stream.Dictionary.Items.TryGetValue("SMask", out var softMaskObj)) {
            return false;
        }

        PdfStream? softMask = ResolveStream(softMaskObj, objects);
        if (softMask is null) {
            return false;
        }

        int softMaskWidth = (int)(softMask.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int softMaskHeight = (int)(softMask.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int softMaskBitsPerComponent = (int)(softMask.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string softMaskColorSpace = GetNameOrEmpty(softMask.Dictionary.Items.TryGetValue("ColorSpace", out var softMaskColorSpaceObj) ? softMaskColorSpaceObj : null, objects);
        if (softMaskWidth != width ||
            softMaskHeight != height ||
            softMaskBitsPerComponent != 8 ||
            !string.Equals(softMaskColorSpace, "DeviceGray", StringComparison.Ordinal) ||
            !TryReadDecodedStreamBytes(softMask, objects, out var alphaPixels)) {
            return false;
        }

        if (indexedPixels.Length == 0 || rgbPalette.Length == 0 || rgbPalette.Length % 3 != 0) {
            return false;
        }

        long sourceRowLengthLong = ((long)width * bitsPerComponent + 7) / 8;
        long expectedLengthLong = sourceRowLengthLong * height;
        long alphaRowLengthLong = width;
        long expectedAlphaLengthLong = alphaRowLengthLong * height;
        long outputRowLengthLong = (long)width * 4;
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            expectedAlphaLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int alphaRowLength = (int)alphaRowLengthLong;
        int expectedAlphaLength = (int)expectedAlphaLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (indexedPixels.Length < expectedLength || alphaPixels.Length < expectedAlphaLength) {
            return false;
        }

        var alphaDecodeTransform = PdfImageDecodeTransform.CreateColor(softMask.Dictionary, 1, objects);
        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        int paletteEntryCount = rgbPalette.Length / 3;
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            int alphaRow = row * alphaRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int paletteIndex = ReadIndexedPixel(indexedPixels, sourceRow, pixel, bitsPerComponent);
                if (decodeTransform is not null) {
                    paletteIndex = decodeTransform.TransformIndexedSample(paletteIndex, bitsPerComponent, paletteEntryCount - 1);
                }

                if (paletteIndex < 0 || paletteIndex >= paletteEntryCount) {
                    return false;
                }

                int paletteOffset = paletteIndex * 3;
                int outputPixel = outputRow + 1 + pixel * 4;
                scanlines[outputPixel] = rgbPalette[paletteOffset];
                scanlines[outputPixel + 1] = rgbPalette[paletteOffset + 1];
                scanlines[outputPixel + 2] = rgbPalette[paletteOffset + 2];
                scanlines[outputPixel + 3] = TransformColorComponent(alphaPixels[alphaRow + pixel], 0, alphaDecodeTransform);
            }
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

    private static int ReadIndexedPixel(byte[] indexedPixels, int rowOffset, int pixelIndex, int bitsPerComponent) {
        if (bitsPerComponent == 8) {
            return indexedPixels[rowOffset + pixelIndex];
        }

        int bitOffset = pixelIndex * bitsPerComponent;
        int sourceByte = indexedPixels[rowOffset + bitOffset / 8];
        int shift = 8 - bitsPerComponent - (bitOffset % 8);
        int mask = (1 << bitsPerComponent) - 1;
        return (sourceByte >> shift) & mask;
    }

    private static byte ConvertDeviceCmykComponentToRgb(byte colorant, byte black) {
        int ink = colorant + black;
        return (byte)(255 - (ink > 255 ? 255 : ink));
    }

    private static byte TransformColorComponent(byte sample, int componentIndex, PdfImageDecodeTransform? decodeTransform) {
        return decodeTransform is null ? sample : decodeTransform.TransformColorComponent(sample, componentIndex);
    }

    private static string GetNameOrEmpty(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        if (resolved is PdfName name) {
            return name.Name;
        }

        return string.Empty;
    }

    private static PdfStream? ResolveStream(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        return resolved as PdfStream;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        return PdfObjectLookup.Resolve(objects, obj);
    }
}
