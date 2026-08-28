using OfficeIMO.Drawing;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static bool IsStructurallyInspectableImage(
        ImageContext context,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        PdfDictionary image = context.Dictionary;
        if (componentCount < 1 ||
            !TryResolveInteger(image, "Width", objects, maximumObjectDepth, 1, int.MaxValue, out int width) ||
            !TryResolveInteger(image, "Height", objects, maximumObjectDepth, 1, int.MaxValue, out int height) ||
            !TryResolveIntegerFromSet(image, "BitsPerComponent", objects, maximumObjectDepth, 1, 2, 4, 8, 16) ||
            !TryResolveInteger(image, "BitsPerComponent", objects, maximumObjectDepth, 1, 16, out int bitsPerComponent) ||
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
                    context.Stream.Data,
                    out int jpegWidth,
                    out int jpegHeight,
                    out int jpegPrecision,
                    out int jpegComponents) &&
                jpegPrecision == bitsPerComponent && jpegComponents == componentCount &&
                OfficeRasterContainerInspector.TryInspect(context.Stream.Data, out OfficeRasterContainerInfo? container) &&
                container != null && container.Format == OfficeImageFormat.Jpeg &&
                jpegWidth == width && jpegHeight == height &&
                container.CanvasWidth == width && container.CanvasHeight == height;
        }

        if (filterName != null && string.Equals(filterName, "JPXDecode", StringComparison.Ordinal)) return false;
        if (!StreamDecoder.TryDecode(
                image,
                context.Stream.Data,
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

    private static bool TryResolveSingleFilterName(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out string? filterName) {
        filterName = null;
        if (!dictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ||
            ResolveObject(objects, filterObject, 0, maximumObjectDepth) is PdfNull) return true;
        if (ResolveObject(objects, filterObject, 0, maximumObjectDepth) is not PdfName name) return false;
        filterName = name.Name;
        return true;
    }

    private static bool HasValidDctDecodeParameters(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parametersObject) ||
            ResolveObject(objects, parametersObject, 0, maximumObjectDepth) is PdfNull) return true;
        if (ResolveObject(objects, parametersObject, 0, maximumObjectDepth) is not PdfDictionary parameters) return false;
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
