using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaxColorSpaceNesting = 8;
    private const int MaxDeviceNComponents = 32;

    private bool TryReadExtendedColorSpaceResource(PdfObject? value, int depth, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (depth > MaxColorSpaceNesting) return false;

        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName directName) {
            return TryReadStandardColorSpaceName(directName.Name, out colorSpace);
        }

        if (resolved is not PdfArray { Items.Count: > 0 } array ||
            ResolveObject(array.Items[0]) is not PdfName arrayName) {
            return false;
        }

        switch (arrayName.Name) {
            case "ICCBased":
                return TryReadIccColorSpace(array, depth, out colorSpace);
            case "Indexed":
            case "I":
                return TryReadIndexedColorSpace(array, depth, out colorSpace);
            case "Separation":
                return TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.Separation, 1, depth, out colorSpace);
            case "DeviceN":
            case "NChannel":
                int componentCount = TryReadDeviceNComponentCount(array);
                return componentCount > 0 &&
                    TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.DeviceN, componentCount, depth, out colorSpace);
            case "CalRGB":
                return array.Items.Count > 1 &&
                    ResolveDictionary(array.Items[1]) is PdfDictionary calibration &&
                    TryReadCalRgbColorSpace(calibration, out colorSpace);
            default:
                return TryReadStandardColorSpaceName(arrayName.Name, out colorSpace);
        }
    }

    private bool TryReadIccColorSpace(PdfArray array, int depth, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 2) return false;
        PdfObject? resolvedProfile = ResolveObject(array.Items[1]);
        PdfDictionary? profile = resolvedProfile switch {
            PdfStream stream => stream.Dictionary,
            PdfDictionary dictionary => dictionary,
            _ => null
        };
        int? components = profile == null
            ? null
            : TryReadInteger(profile.Items.TryGetValue("N", out PdfObject? count) ? count : null);
        PdfPageColorSpaceKind kind = components switch {
            1 => PdfPageColorSpaceKind.DeviceGray,
            3 => PdfPageColorSpaceKind.DeviceRgb,
            4 => PdfPageColorSpaceKind.DeviceCmyk,
            _ => PdfPageColorSpaceKind.Pattern
        };
        if (kind == PdfPageColorSpaceKind.Pattern) return false;
        IReadOnlyList<double>? ranges = null;
        if (profile != null && profile.Items.TryGetValue("Range", out PdfObject? rangeObject)) {
            ranges = ReadNumberArray(rangeObject);
            if (ranges.Count != components * 2) return false;
            for (int index = 0; index < components; index++) {
                double minimum = ranges[index * 2];
                double maximum = ranges[index * 2 + 1];
                if (!IsFinite(minimum) || !IsFinite(maximum) || minimum >= maximum) return false;
            }
        }

        if (resolvedProfile is PdfStream profileStream) {
            if (PdfIccProfileCache.TryRead(profileStream, _objects, _limits.MaxDecodedStreamBytes, out OfficeIccColorProfile? parsedProfile) &&
                parsedProfile != null && parsedProfile.ComponentCount == components) {
                colorSpace = PdfPageColorSpace.IccBased(parsedProfile, ranges);
                return true;
            }
        }

        if (profile != null &&
            profile.Items.TryGetValue("Alternate", out PdfObject? alternateObject) &&
            TryReadExtendedColorSpaceResource(alternateObject, depth + 1, out PdfPageColorSpace alternate) &&
            alternate.Kind is not PdfPageColorSpaceKind.Pattern and not PdfPageColorSpaceKind.Indexed &&
            alternate.ComponentCount == components) {
            colorSpace = PdfPageColorSpace.IccFallback(alternate);
            return true;
        }

        colorSpace = PdfPageColorSpace.IccBased(kind);
        return true;
    }

    private bool TryReadIndexedColorSpace(PdfArray array, int depth, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 ||
            !TryReadExtendedColorSpaceResource(array.Items[1], depth + 1, out PdfPageColorSpace baseColorSpace) ||
            baseColorSpace.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            TryReadInteger(array.Items[2]) is not int highValue ||
            highValue < 0 || highValue > 255) {
            return false;
        }

        int componentCount = baseColorSpace.ComponentCount;
        int paletteCount = highValue + 1;
        int lookupLength = checked(paletteCount * componentCount);
        int lookupLimit = Math.Min(lookupLength, _limits.MaxDecodedStreamBytes);
        if (!PdfIndexedImageNormalizer.TryReadIndexedLookupBytes(
                array.Items[3],
                _objects,
                lookupLimit,
                out byte[] lookupBytes)) return false;
        if (lookupBytes.Length < paletteCount * componentCount) return false;

        var palette = new OfficeColor[paletteCount];
        var components = new double[componentCount];
        for (int entry = 0; entry < paletteCount; entry++) {
            for (int component = 0; component < componentCount; component++) {
                components[component] = lookupBytes[entry * componentCount + component] / 255D;
            }
            if (!baseColorSpace.TryConvertColor(components, out palette[entry])) return false;
        }

        colorSpace = PdfPageColorSpace.Indexed(palette, baseColorSpace.UsesIccApproximation);
        return true;
    }

    private bool TryReadAlternateColorSpace(
        PdfArray array,
        PdfPageColorSpaceKind kind,
        int componentCount,
        int depth,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 || componentCount < 1 || componentCount > MaxDeviceNComponents ||
            !TryReadExtendedColorSpaceResource(array.Items[2], depth + 1, out PdfPageColorSpace alternate) ||
            alternate.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            !PdfColorSpaceFunctionResolver.TryCreateTintTransform(
                array.Items[3],
                componentCount,
                alternate.ComponentCount,
                _objects,
                _limits.MaxDecodedStreamBytes,
                out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform)) {
            return false;
        }

        colorSpace = PdfPageColorSpace.Alternate(kind, componentCount, alternate, transform);
        return true;
    }

    private int TryReadDeviceNComponentCount(PdfArray array) {
        if (array.Items.Count < 2 || ResolveObject(array.Items[1]) is not PdfArray names) return 0;
        if (names.Items.Count < 1 || names.Items.Count > MaxDeviceNComponents) return 0;
        return names.Items.All(item => ResolveObject(item) is PdfName) ? names.Items.Count : 0;
    }

}
