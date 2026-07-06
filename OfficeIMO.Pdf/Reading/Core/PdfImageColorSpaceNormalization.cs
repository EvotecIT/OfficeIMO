namespace OfficeIMO.Pdf;

internal sealed class PdfImageColorSpaceNormalization {
    private PdfImageColorSpaceNormalization(int sourceColorCount, int pngColorType) {
        SourceColorCount = sourceColorCount;
        PngColorType = pngColorType;
    }

    internal int SourceColorCount { get; }

    internal int PngColorType { get; }

    internal static bool TryResolve(
        PdfObject? colorSpaceObj,
        string colorSpaceName,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorSpaceNormalization normalization) {
        if (TryCreateFromDeviceColorSpace(colorSpaceName, out normalization)) {
            return true;
        }

        if (ResolveObject(colorSpaceObj, objects) is not PdfArray colorSpaceArray ||
            colorSpaceArray.Items.Count < 2 ||
            ResolveObject(colorSpaceArray.Items[0], objects) is not PdfName colorSpaceKind ||
            (!string.Equals(colorSpaceKind.Name, "ICCBased", StringComparison.Ordinal) &&
             !string.Equals(colorSpaceKind.Name, "ICC", StringComparison.Ordinal)) ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfStream profileStream) {
            normalization = null!;
            return false;
        }

        int? componentCount = TryReadIccComponentCount(profileStream, objects);
        if (!componentCount.HasValue &&
            profileStream.Dictionary.Items.TryGetValue("Alternate", out var alternateObj) &&
            ResolveObject(alternateObj, objects) is PdfName alternateName &&
            TryCreateFromDeviceColorSpace(alternateName.Name, out normalization)) {
            return true;
        }

        if (!componentCount.HasValue) {
            normalization = null!;
            return false;
        }

        switch (componentCount.Value) {
            case 1:
                normalization = new PdfImageColorSpaceNormalization(1, 0);
                return true;
            case 3:
                normalization = new PdfImageColorSpaceNormalization(3, 2);
                return true;
            case 4:
                normalization = new PdfImageColorSpaceNormalization(4, 2);
                return true;
            default:
                normalization = null!;
                return false;
        }
    }

    private static bool TryCreateFromDeviceColorSpace(string colorSpaceName, out PdfImageColorSpaceNormalization normalization) {
        if (string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal)) {
            normalization = new PdfImageColorSpaceNormalization(1, 0);
            return true;
        }

        if (string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal)) {
            normalization = new PdfImageColorSpaceNormalization(3, 2);
            return true;
        }

        if (string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal)) {
            normalization = new PdfImageColorSpaceNormalization(4, 2);
            return true;
        }

        normalization = null!;
        return false;
    }

    private static int? TryReadIccComponentCount(PdfStream profileStream, Dictionary<int, PdfIndirectObject> objects) {
        if (!profileStream.Dictionary.Items.TryGetValue("N", out var countObject) ||
            ResolveObject(countObject, objects) is not PdfNumber countNumber ||
            countNumber.Value < 0 ||
            countNumber.Value > int.MaxValue ||
            System.Math.Truncate(countNumber.Value) != countNumber.Value) {
            return null;
        }

        return (int)countNumber.Value;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        return PdfObjectLookup.Resolve(objects, obj);
    }
}
