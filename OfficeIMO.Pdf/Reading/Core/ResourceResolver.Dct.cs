using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static bool IsDctFilterChain(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects) =>
        TryGetDctFilters(imageDictionary, objects, out _);

    private static bool HasDctPrefixFilters(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects) =>
        TryGetDctFilters(imageDictionary, objects, out List<PdfName>? filters) && filters.Count > 1;

    private static bool HasTrailingDctFilter(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        PdfObject? filterObject = imageDictionary.Items.TryGetValue("Filter", out PdfObject? fullName)
            ? fullName
            : imageDictionary.Items.TryGetValue("F", out PdfObject? abbreviatedName)
                ? abbreviatedName
                : null;
        PdfObject? resolvedFilter = PdfObjectLookup.ResolveChain(objects, filterObject);
        if (resolvedFilter is PdfName singleFilter) return IsDctFilterName(singleFilter.Name);
        if (resolvedFilter is not PdfArray { Items.Count: > 0 } filterArray) return false;
        return PdfObjectLookup.ResolveChain(objects, filterArray.Items[filterArray.Items.Count - 1]) is PdfName lastFilter &&
               IsDctFilterName(lastFilter.Name);
    }

    private static bool RequiresDctColorNormalization(
        PdfDictionary imageDictionary,
        string colorSpace,
        string? transparencyMaskKind,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!string.IsNullOrEmpty(transparencyMaskKind)) {
            return true;
        }
        if (!PdfImageDecodeTransform.TryCreateColorDeclaration(
                imageDictionary,
                GetDeclaredDeviceColorCount(colorSpace),
                objects,
                out PdfImageDecodeTransform? decodeTransform) ||
            decodeTransform is not null) {
            return true;
        }
        if (!TryReadDctColorTransform(imageDictionary, objects, out _, out bool hasAuthoredColorTransform) ||
            hasAuthoredColorTransform) return true;

        return colorSpace is not ("DeviceGray" or "G" or "DeviceRGB" or "RGB" or "DeviceCMYK" or "CMYK");
    }

    private static bool CanPassThroughDctPayload(
        PdfDictionary imageDictionary,
        string colorSpace,
        Dictionary<int, PdfIndirectObject> objects) {
        if (colorSpace is not ("DeviceGray" or "G" or "DeviceRGB" or "RGB" or "DeviceCMYK" or "CMYK")) {
            return false;
        }

        int componentCount = GetDeclaredDeviceColorCount(colorSpace);
        if (!PdfImageDecodeTransform.TryCreateColorDeclaration(
                imageDictionary,
                componentCount,
                objects,
                out PdfImageDecodeTransform? decodeTransform) ||
            decodeTransform is not null ||
            !PdfImageColorKeyMask.TryCreateDeclaration(
                imageDictionary,
                componentCount,
                bitsPerComponent: 8,
                objects,
                out _)) return false;

        return TryReadDctColorTransform(imageDictionary, objects, out _, out bool hasAuthoredColorTransform) &&
               !hasAuthoredColorTransform;
    }

    private static bool TryGetDctPayload(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out byte[] jpegBytes) {
        jpegBytes = Array.Empty<byte>();
        if (!TryGetDctFilters(stream.Dictionary, objects, out List<PdfName>? filters)) return false;
        if (filters.Count == 1) {
            jpegBytes = stream.Data;
            return true;
        }

        var prefixDictionary = new PdfDictionary();
        var prefixFilters = new PdfArray();
        for (int index = 0; index < filters.Count - 1; index++) {
            prefixFilters.Items.Add(filters[index]);
        }
        prefixDictionary.Items["Filter"] = prefixFilters;

        PdfObject? decodeParmsObject = stream.Dictionary.Items.TryGetValue("DecodeParms", out PdfObject? fullName)
            ? fullName
            : stream.Dictionary.Items.TryGetValue("DP", out PdfObject? abbreviatedName)
                ? abbreviatedName
                : null;
        PdfObject? resolvedDecodeParms = PdfObjectLookup.ResolveChain(objects, decodeParmsObject);
        if (resolvedDecodeParms is not null and not PdfNull) {
            if (resolvedDecodeParms is not PdfArray decodeParmsArray || decodeParmsArray.Items.Count != filters.Count) {
                return false;
            }

            var prefixDecodeParms = new PdfArray();
            for (int index = 0; index < filters.Count - 1; index++) {
                prefixDecodeParms.Items.Add(decodeParmsArray.Items[index]);
            }
            prefixDictionary.Items["DecodeParms"] = prefixDecodeParms;
        }

        try {
            jpegBytes = Filters.StreamDecoder.DecodeRequired(
                prefixDictionary,
                stream.Data,
                objects,
                maxDecodedStreamBytes);
            return OfficeJpegCodec.IsJpeg(jpegBytes);
        } catch (PdfReadLimitException) {
            throw;
        } catch (InvalidDataException) {
            jpegBytes = Array.Empty<byte>();
            return false;
        }
    }

    private static bool TryDecodeDctImage(
        PdfStream stream,
        int width,
        int height,
        int expectedColorCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out byte[] pixels) {
        pixels = Array.Empty<byte>();
        if (expectedColorCount < 1 || expectedColorCount > 4 || maxDecodedStreamBytes <= 0) {
            return false;
        }

        long expectedDecodedBytes = (long)width * height * expectedColorCount;
        if (expectedDecodedBytes > maxDecodedStreamBytes) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.DecodedStreamBytes,
                maxDecodedStreamBytes,
                expectedDecodedBytes);
        }

        if (!TryReadDctColorTransform(stream.Dictionary, objects, out int? requestedColorTransform, out _) ||
            !TryGetDctPayload(stream, objects, maxDecodedStreamBytes, out byte[] jpegBytes) ||
            !HasExpectedDctFrame(jpegBytes, width, height, expectedColorCount) ||
            !OfficeJpegCodec.TryDecodeColorComponents(
                jpegBytes,
                requestedColorTransform,
                usePdfColorTransformDefault: true,
                out pixels,
                out int decodedWidth,
                out int decodedHeight,
                out int decodedColorCount) ||
            decodedWidth != width ||
            decodedHeight != height ||
            decodedColorCount != expectedColorCount ||
            pixels.LongLength != expectedDecodedBytes) {
            pixels = Array.Empty<byte>();
            return false;
        }

        return true;
    }

    internal static bool HasExpectedDctFrame(byte[] jpegBytes, int width, int height, int expectedColorCount) =>
        OfficeImageReader.TryIdentify(jpegBytes, null, out OfficeImageInfo jpegInfo) &&
        jpegInfo.Format == OfficeImageFormat.Jpeg &&
        jpegInfo.Width == width &&
        jpegInfo.Height == height &&
        PdfWriter.TryGetJpegFrameMetadata(jpegBytes, out int frameColorCount, out _) &&
        frameColorCount == expectedColorCount;

    private static bool TryReadDctColorTransform(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out int? colorTransform,
        out bool hasAuthoredColorTransform) {
        colorTransform = null;
        hasAuthoredColorTransform = false;
        if (!TryGetDctDecodeParms(imageDictionary, objects, out PdfDictionary? decodeParms)) return false;
        if (decodeParms is null) return true;
        if (!decodeParms.Items.TryGetValue("ColorTransform", out PdfObject? colorTransformObject)) return true;

        PdfObject? resolvedColorTransform = PdfObjectLookup.ResolveChain(objects, colorTransformObject);
        if (resolvedColorTransform is null or PdfNull) return true;
        hasAuthoredColorTransform = true;
        if (resolvedColorTransform is not PdfNumber number ||
            number.Value != Math.Truncate(number.Value) ||
            number.Value is not (0D or 1D)) {
            return false;
        }

        colorTransform = (int)number.Value;
        return true;
    }

    private static int GetDeclaredDeviceColorCount(string colorSpace) {
        switch (colorSpace) {
            case "DeviceGray":
            case "G":
                return 1;
            case "DeviceRGB":
            case "RGB":
                return 3;
            case "DeviceCMYK":
            case "CMYK":
                return 4;
            default:
                return 0;
        }
    }

    private static bool TryGetDctDecodeParms(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfDictionary? decodeParms) {
        decodeParms = null;
        if (!TryGetDctFilters(imageDictionary, objects, out List<PdfName>? filters)) return false;
        PdfObject? decodeParmsObject = imageDictionary.Items.TryGetValue("DecodeParms", out PdfObject? fullName)
            ? fullName
            : imageDictionary.Items.TryGetValue("DP", out PdfObject? abbreviatedName)
                ? abbreviatedName
                : null;
        PdfObject? resolvedDecodeParms = PdfObjectLookup.ResolveChain(objects, decodeParmsObject);
        if (resolvedDecodeParms is null or PdfNull) return true;
        if (filters.Count == 1) {
            PdfObject? filterObject = imageDictionary.Items.TryGetValue("Filter", out PdfObject? fullFilterName)
                ? fullFilterName
                : imageDictionary.Items.TryGetValue("F", out PdfObject? abbreviatedFilterName)
                    ? abbreviatedFilterName
                    : null;
            bool filterWasAuthoredAsArray = PdfObjectLookup.ResolveChain(objects, filterObject) is PdfArray;
            if (filterWasAuthoredAsArray) {
                if (resolvedDecodeParms is not PdfArray { Items.Count: 1 } singleDecodeParmsArray) return false;
                PdfObject? resolvedSingleDecodeParms = PdfObjectLookup.ResolveChain(objects, singleDecodeParmsArray.Items[0]);
                if (resolvedSingleDecodeParms is null or PdfNull) return true;
                decodeParms = resolvedSingleDecodeParms as PdfDictionary;
                return decodeParms is not null;
            }

            decodeParms = resolvedDecodeParms as PdfDictionary;
            return decodeParms is not null;
        }
        if (resolvedDecodeParms is not PdfArray decodeParmsArray || decodeParmsArray.Items.Count != filters.Count) {
            return false;
        }

        PdfObject? resolvedDctParms = PdfObjectLookup.ResolveChain(objects, decodeParmsArray.Items[filters.Count - 1]);
        if (resolvedDctParms is null or PdfNull) return true;
        decodeParms = resolvedDctParms as PdfDictionary;
        return decodeParms is not null;
    }

    private static bool TryGetDctFilters(
        PdfDictionary imageDictionary,
        Dictionary<int, PdfIndirectObject> objects,
        out List<PdfName> filters) {
        filters = new List<PdfName>();
        PdfObject? filterObject = imageDictionary.Items.TryGetValue("Filter", out PdfObject? fullName)
            ? fullName
            : imageDictionary.Items.TryGetValue("F", out PdfObject? abbreviatedName)
                ? abbreviatedName
                : null;
        PdfObject? resolvedFilter = PdfObjectLookup.ResolveChain(objects, filterObject);
        if (resolvedFilter is PdfName singleFilter) {
            filters.Add(singleFilter);
        } else if (resolvedFilter is PdfArray filterArray) {
            for (int index = 0; index < filterArray.Items.Count; index++) {
                if (PdfObjectLookup.ResolveChain(objects, filterArray.Items[index]) is not PdfName filterName) return false;
                filters.Add(filterName);
            }
        } else {
            return false;
        }

        if (filters.Count == 0 || !IsDctFilterName(filters[filters.Count - 1].Name)) return false;
        for (int index = 0; index < filters.Count - 1; index++) {
            if (IsDctFilterName(filters[index].Name) || !Filters.StreamDecoder.IsSupportedFilter(filters[index].Name)) {
                return false;
            }
        }
        return true;
    }

    private static bool IsDctFilterName(string name) =>
        string.Equals(name, "DCTDecode", StringComparison.Ordinal) ||
        string.Equals(name, "DCT", StringComparison.Ordinal);
}
