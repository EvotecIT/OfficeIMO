using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static bool IsStructurallyInspectableSoftMask(
        ImageContext owner,
        PdfStream softMask,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        PdfDictionary dictionary = softMask.Dictionary;
        if (dictionary.Items.ContainsKey("Mask") || dictionary.Items.ContainsKey("SMask")) return false;

        PdfObject? softMaskColorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpace)
            ? colorSpace
            : null;
        ColorSpaceUsage softMaskUsage = ClassifyColorSpace(
            softMaskColorSpace,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes,
            owner.Aliases);
        if (!softMaskUsage.IsKnown || softMaskUsage.ComponentCount != 1 || softMaskUsage.UsesPattern ||
            !IsStructurallyInspectableImage(
                dictionary,
                softMask.Data,
                componentCount: 1,
                objects,
                maximumObjectDepth,
                maximumDecodedStreamBytes)) {
            return false;
        }

        if (!dictionary.Items.TryGetValue("Matte", out PdfObject? matteObject) ||
            ResolveObject(objects, matteObject, 0, maximumObjectDepth) is PdfNull) return true;
        PdfObject? ownerColorSpace = owner.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? ownerColorSpaceObject)
            ? ownerColorSpaceObject
            : null;
        ColorSpaceUsage ownerUsage = ClassifyColorSpace(
            ownerColorSpace,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes,
            owner.Aliases);
        return ownerUsage.IsKnown && ownerUsage.ComponentCount > 0 &&
            HasExactFiniteNumberArray(
                dictionary,
                "Matte",
                ownerUsage.ComponentCount,
                objects,
                maximumObjectDepth,
                out _);
    }

    private static bool HasStructurallyValidExplicitMask(
        ImageContext owner,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        bool hasSoftMask = owner.Dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject) &&
            ResolveObject(objects, softMaskObject, 0, maximumObjectDepth) is not PdfNull and not PdfName { Name: "None" };
        if (!owner.Dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject) ||
            ResolveObject(objects, maskObject, 0, maximumObjectDepth) is PdfNull) return true;
        if (hasSoftMask || componentCount < 1) return false;

        PdfObject? resolvedMask = ResolveObject(objects, maskObject, 0, maximumObjectDepth);
        if (resolvedMask is PdfArray colorKeyMask) {
            if (colorKeyMask.Items.Count != checked(componentCount * 2) ||
                !TryResolveInteger(owner.Dictionary, "BitsPerComponent", objects, maximumObjectDepth, 1, 16, out int bitsPerComponent)) {
                return false;
            }
            int maximumSample = (1 << bitsPerComponent) - 1;
            for (int component = 0; component < componentCount; component++) {
                if (!TryResolveBoundedInteger(colorKeyMask.Items[component * 2], objects, maximumObjectDepth, 0, maximumSample, out int minimum) ||
                    !TryResolveBoundedInteger(colorKeyMask.Items[component * 2 + 1], objects, maximumObjectDepth, minimum, maximumSample, out _)) {
                    return false;
                }
            }
            return true;
        }

        if (resolvedMask is not PdfStream maskStream ||
            !string.Equals(
                ResolveName(
                    maskStream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
                    objects,
                    maximumObjectDepth),
                "Image",
                StringComparison.Ordinal) ||
            maskStream.Dictionary.Items.ContainsKey("Mask") ||
            maskStream.Dictionary.Items.ContainsKey("SMask") ||
            ResolveObject(
                objects,
                maskStream.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) ? imageMaskObject : null,
                0,
                maximumObjectDepth) is not PdfBoolean { Value: true }) {
            return false;
        }
        return IsStructurallyInspectableImageMask(
            maskStream.Dictionary,
            maskStream.Data,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes);
    }

    private static bool IsStructurallyInspectableImage(
        ImageContext context,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) => IsStructurallyInspectableImage(
            context.Dictionary,
            context.Stream.Data,
            componentCount,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes);

    private static bool IsStructurallyInspectableImage(
        PdfDictionary image,
        byte[] data,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (componentCount < 1 ||
            !TryResolveInteger(image, "Width", objects, maximumObjectDepth, 1, int.MaxValue, out int width) ||
            !TryResolveInteger(image, "Height", objects, maximumObjectDepth, 1, int.MaxValue, out int height) ||
            !TryResolveIntegerFromSet(image, "BitsPerComponent", objects, maximumObjectDepth, 1, 2, 4, 8, 16) ||
            !TryResolveInteger(image, "BitsPerComponent", objects, maximumObjectDepth, 1, 16, out int bitsPerComponent) ||
            !HasOptionalBoolean(image, "ImageMask", objects, maximumObjectDepth) ||
            !HasOptionalBoolean(image, "Interpolate", objects, maximumObjectDepth) ||
            !HasOptionalExactFiniteNumberArray(
                image,
                "Decode",
                checked(componentCount * 2),
                objects,
                maximumObjectDepth)) {
            return false;
        }

        if (TryResolveSingleFilterName(image, objects, maximumObjectDepth, out string? filterName) &&
            string.Equals(filterName, "DCTDecode", StringComparison.Ordinal)) {
            return HasValidDctDecodeParameters(image, objects, maximumObjectDepth) &&
                TryReadJpegFrame(
                    data,
                    out int jpegWidth,
                    out int jpegHeight,
                    out int jpegPrecision,
                    out int jpegComponents) &&
                jpegPrecision == bitsPerComponent && jpegComponents == componentCount &&
                OfficeRasterContainerInspector.TryInspect(data, out OfficeRasterContainerInfo? container) &&
                container != null && container.Format == OfficeImageFormat.Jpeg &&
                jpegWidth == width && jpegHeight == height &&
                container.CanvasWidth == width && container.CanvasHeight == height;
        }

        if (filterName != null && string.Equals(filterName, "JPXDecode", StringComparison.Ordinal)) return false;
        if (!StreamDecoder.TryDecode(
                image,
                data,
                maximumDecodedStreamBytes,
                out byte[] decoded,
                objects)) return false;

        long rowBits;
        long rowBytes;
        long requiredBytes;
        try {
            rowBits = checked((long)width * componentCount * bitsPerComponent);
            rowBytes = checked((rowBits + 7L) / 8L);
            requiredBytes = checked(rowBytes * height);
        } catch (OverflowException) {
            return false;
        }
        return requiredBytes <= maximumDecodedStreamBytes && decoded.LongLength == requiredBytes;
    }

    private static bool IsStructurallyInspectableImageMask(
        ImageContext context,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) => IsStructurallyInspectableImageMask(
            context.Dictionary,
            context.Stream.Data,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes);

    private static bool IsStructurallyInspectableImageMask(
        PdfDictionary image,
        byte[] data,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (!TryResolveInteger(image, "Width", objects, maximumObjectDepth, 1, int.MaxValue, out int width) ||
            !TryResolveInteger(image, "Height", objects, maximumObjectDepth, 1, int.MaxValue, out int height) ||
            image.Items.ContainsKey("ColorSpace") ||
            (image.Items.TryGetValue("BitsPerComponent", out _) &&
                !TryResolveInteger(image, "BitsPerComponent", objects, maximumObjectDepth, 1, 1, out _)) ||
            !HasOptionalBoolean(image, "Interpolate", objects, maximumObjectDepth) ||
            !HasOptionalExactFiniteNumberArray(image, "Decode", 2, objects, maximumObjectDepth) ||
            !TryResolveSingleFilterName(image, objects, maximumObjectDepth, out string? filterName) ||
            string.Equals(filterName, "DCTDecode", StringComparison.Ordinal) ||
            string.Equals(filterName, "JPXDecode", StringComparison.Ordinal) ||
            !StreamDecoder.TryDecode(image, data, maximumDecodedStreamBytes, out byte[] decoded, objects)) {
            return false;
        }

        long requiredBytes;
        try {
            requiredBytes = checked(((long)width + 7L) / 8L * height);
        } catch (OverflowException) {
            return false;
        }
        return requiredBytes <= maximumDecodedStreamBytes && decoded.LongLength == requiredBytes;
    }

    private static bool TryResolveSingleFilterName(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out string? filterName) {
        filterName = null;
        if (!dictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ||
            ResolveObject(objects, filterObject, 0, maximumObjectDepth) is PdfNull) return true;
        PdfObject? resolved = ResolveObject(objects, filterObject, 0, maximumObjectDepth);
        if (resolved is PdfArray filters) {
            if (filters.Items.Count != 1) return false;
            resolved = ResolveObject(objects, filters.Items[0], 0, maximumObjectDepth);
        }
        if (resolved is not PdfName name) return false;
        filterName = name.Name;
        return true;
    }

    private static bool HasValidDctDecodeParameters(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parametersObject) ||
            ResolveObject(objects, parametersObject, 0, maximumObjectDepth) is PdfNull) return true;
        PdfObject? resolved = ResolveObject(objects, parametersObject, 0, maximumObjectDepth);
        if (resolved is PdfArray parameterSets) {
            if (parameterSets.Items.Count != 1) return false;
            resolved = ResolveObject(objects, parameterSets.Items[0], 0, maximumObjectDepth);
            if (resolved is PdfNull) return true;
        }
        if (resolved is not PdfDictionary parameters) return false;
        return !parameters.Items.ContainsKey("ColorTransform") ||
            TryResolveInteger(parameters, "ColorTransform", objects, maximumObjectDepth, 0, 1, out _);
    }

    private static bool TryReadJpegFrame(
        byte[] data,
        out int width,
        out int height,
        out int precision,
        out int components) {
        width = 0;
        height = 0;
        precision = 0;
        components = 0;
        if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8) return false;
        int offset = 2;
        while (offset < data.Length) {
            if (data[offset++] != 0xFF) return false;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) return false;
            byte marker = data[offset++];
            if (marker == 0xD9 || marker == 0xDA) return false;
            if (marker is >= 0xD0 and <= 0xD8 || marker == 0x01) continue;
            if (offset > data.Length - 2) return false;
            int length = (data[offset] << 8) | data[offset + 1];
            if (length < 2 || offset > data.Length - length) return false;
            if (IsJpegStartOfFrame(marker)) {
                if (length < 8) return false;
                precision = data[offset + 2];
                height = (data[offset + 3] << 8) | data[offset + 4];
                width = (data[offset + 5] << 8) | data[offset + 6];
                components = data[offset + 7];
                return width > 0 && height > 0 && components is >= 1 and <= 4 &&
                    length == 8 + components * 3;
            }
            offset += length;
        }
        return false;
    }

    private static bool IsJpegStartOfFrame(byte marker) =>
        marker is >= 0xC0 and <= 0xC3 or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB or >= 0xCD and <= 0xCF;
}
